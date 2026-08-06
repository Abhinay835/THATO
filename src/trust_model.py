"""Adaptive XGBoost trust model used by TAHTO experiments.

The model is deliberately evaluated on temporally later telemetry, never on
the synthetic samples used to create its current training window.
"""
from __future__ import annotations

import numpy as np
from xgboost import XGBClassifier
from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score
from sklearn.model_selection import StratifiedKFold
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import MinMaxScaler


class TrustModel:
    def __init__(self, random_state: int = 42):
        self.random_state = random_state
        self.model = None
        self.trained = False
        self.training_samples = 0
        self.update_count = 0
        self.accuracy = self.precision = self.recall = self.f1 = 0.0
        self.cv_scores = {}

    def _classifier(self):
        return XGBClassifier(
            n_estimators=80, max_depth=3, learning_rate=0.08,
            subsample=0.9, colsample_bytree=0.9, eval_metric="logloss",
            random_state=self.random_state, n_jobs=1,
        )

    def fit(self, X, y, update=False):
        """Fit from labelled telemetry observed before the current decision window."""
        X, y = np.asarray(X, dtype=float), np.asarray(y, dtype=int)
        if len(np.unique(y)) < 2:
            raise ValueError("Trust training requires both trusted and untrusted examples")
        self.model = Pipeline([("scale", MinMaxScaler()), ("xgb", self._classifier())])
        self.model.fit(X, y)
        self.trained = True
        self.training_samples = len(y)
        if update:
            self.update_count += 1

    def cross_validate(self, X, y, folds=5):
        """Leakage-free CV: each fold fits its own scaler and model."""
        X, y = np.asarray(X, dtype=float), np.asarray(y, dtype=int)
        n_splits = min(folds, int(np.bincount(y).min()))
        if n_splits < 2:
            return {}
        metrics = {"accuracy": [], "precision": [], "recall": [], "f1": []}
        for train, test in StratifiedKFold(n_splits=n_splits, shuffle=True,
                                           random_state=self.random_state).split(X, y):
            model = Pipeline([("scale", MinMaxScaler()), ("xgb", self._classifier())])
            model.fit(X[train], y[train])
            pred = model.predict(X[test])
            metrics["accuracy"].append(accuracy_score(y[test], pred))
            metrics["precision"].append(precision_score(y[test], pred, zero_division=0))
            metrics["recall"].append(recall_score(y[test], pred, zero_division=0))
            metrics["f1"].append(f1_score(y[test], pred, zero_division=0))
        self.cv_scores = {f"cv_{name}_mean": round(float(np.mean(values)), 4)
                          for name, values in metrics.items()}
        self.cv_scores.update({f"cv_{name}_std": round(float(np.std(values)), 4)
                               for name, values in metrics.items()})
        return self.cv_scores

    def evaluate(self, X, y):
        """Evaluate the currently fitted model on an unseen, later telemetry window."""
        pred = self.model.predict(np.asarray(X, dtype=float))
        y = np.asarray(y, dtype=int)
        self.accuracy = float(accuracy_score(y, pred))
        self.precision = float(precision_score(y, pred, zero_division=0))
        self.recall = float(recall_score(y, pred, zero_division=0))
        self.f1 = float(f1_score(y, pred, zero_division=0))
        return {"accuracy": self.accuracy, "precision": self.precision,
                "recall": self.recall, "f1": self.f1}

    def predict_trust(self, feature_vector):
        if not self.trained:
            return 0.5
        return float(self.model.predict_proba(np.asarray(feature_vector, dtype=float).reshape(1, -1))[0][1])
