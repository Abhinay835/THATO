"""Fair, trace-driven multi-seed evaluation for TAHTO."""
from __future__ import annotations
import json, time
from pathlib import Path
import numpy as np
from simulator import EdgeNode, generate_workload, choose_tahto, choose_performance, choose_static_reputation, CLOUD_LATENCY
from trust_model import TrustModel

NUM_TASKS, WARMUP_WINDOWS, UPDATE_EVERY = 500, 8, 50
SCENARIOS = ("S1", "S2", "S3", "S4", "S5", "S6")
SCHEDULERS = ("TAHTO", "B1", "B2", "B3")
DEFAULT_SEEDS = (42, 123, 2026, 7, 19, 77, 101, 314, 808, 999)

def make_nodes(scenario, seed):
    configs = {
        "S1": [(2.0,100,False,1),(1.8,100,False,1),(2.2,100,False,1),(1.5,100,False,1),(2.0,100,False,1)],
        "S2": [(2.0,20,False,1),(1.8,20,False,1),(2.2,20,False,1),(1.5,20,False,1),(2.0,20,False,1)],
        "S3": [(2.0,1,False,1),(1.8,1,False,1),(2.2,1,False,1),(1.5,1,False,1),(2.0,1,False,1)],
        "S4": [(2.0,100,False,0),(1.8,100,False,0),(2.2,100,False,1),(1.5,100,False,1),(2.0,100,False,1)],
        "S5": [(2.0,100,False,1),(1.8,100,False,1),(2.2,100,False,1),(1.5,100,True,1),(2.0,100,True,1)],
        "S6": [(2.0,5,False,0),(1.8,5,False,0),(2.2,5,False,1),(1.5,5,True,1),(2.0,5,True,1)],
    }[scenario]
    # In S5/S6, attackers behave normally until 20 seconds, then compromise
    # activates. Static reputation is therefore historical rather than prophetic.
    attack_start = 20_000.0 if scenario in {"S5", "S6"} else 0.0
    return [EdgeNode(i, *config, attack_start_ms=attack_start, seed=seed)
            for i, config in enumerate(configs)]

def warmup_model(nodes, seed):
    # A fixed, separate historical calibration period supplies both classes even
    # for benign scenarios. It is never used as the evaluation window.
    history = [EdgeNode(i, 2.0, 50, malicious=(i >= 3), attack_start_ms=0,
                        seed=seed + 5000)
               for i in range(5)]
    X, y = [], []
    for w in range(WARMUP_WINDOWS):
        for node in history:
            node.update_telemetry(w * 1000)
            X.append(node.features()); y.append(node.label(w * 1000))
    # Historical reputation differentiates normal nodes but is fixed before the
    # S5/S6 compromise event, so it cannot react once attacker behaviour changes.
    rep_rng = np.random.default_rng(seed + 7000)
    for node in nodes:
        node.static_reputation = float(rep_rng.uniform(.72, .96))
    model = TrustModel(seed)
    model.fit(X, y)
    model.cross_validate(X, y)
    return model, X, y

def run_scenario(scenario, scheduler, seed=42, alpha=.30, beta=.20, gamma=.20, delta=.30):
    nodes = make_nodes(scenario, seed)
    workload = generate_workload(NUM_TASKS, seed=seed + 1000)
    tm, train_x, train_y = warmup_model(nodes, seed)
    latencies=[]; energies=[]; overhead=[]; cpu=[]; mal_hits=cloud=completed=0
    stream_x=[]; stream_y=[]; eval_metrics=[]; window_metrics=[]
    for i, task in enumerate(workload):
        if i % UPDATE_EVERY == 0:
            snapshot_x=[]; snapshot_y=[]
            for node in nodes:
                node.update_telemetry(task.arrival_ms)
                snapshot_x.append(node.features()); snapshot_y.append(node.label(task.arrival_ms))
            # Test on the current, unseen window before using it for retraining.
            metrics = tm.evaluate(snapshot_x, snapshot_y)
            eval_metrics.append(metrics)
            window_metrics.append({"time_ms": task.arrival_ms, **metrics})
            stream_x.extend(snapshot_x); stream_y.extend(snapshot_y)
            if i and len(set(train_y + stream_y)) == 2:
                train_x.extend(stream_x); train_y.extend(stream_y)
                tm.fit(train_x, train_y, update=True)
                stream_x, stream_y = [], []
            for node in nodes: node.trust_score = tm.predict_trust(node.features())
            active = [n.cpu_util for n in nodes if n.availability >= .1]
            if active: cpu.append(float(np.mean(active)))
        t0=time.perf_counter()
        if scheduler == "TAHTO": dest, lat, eng = choose_tahto(task, nodes, task.arrival_ms, alpha,beta,gamma,delta)
        elif scheduler == "B1": dest, lat, eng = choose_performance(task, nodes, task.arrival_ms)
        elif scheduler == "B2": dest, lat, eng = choose_static_reputation(task, nodes, task.arrival_ms)
        else: dest, lat, eng = "cloud", CLOUD_LATENCY, .05
        overhead.append((time.perf_counter()-t0)*1e6); latencies.append(lat); energies.append(eng)
        if dest == "cloud": cloud += 1; completed += 1
        elif nodes[dest].is_compromised(task.arrival_ms): mal_hits += 1
        else: completed += 1
    mean_ml = {k: float(np.mean([m[k] for m in eval_metrics])) for k in ("accuracy","precision","recall","f1")}
    return {"scenario":scenario,"scheduler":scheduler,"seed":seed,
            "avg_latency_ms":float(np.mean(latencies)),"avg_energy_j":float(np.mean(energies)),
            "task_completion_pct":100*completed/NUM_TASKS,"malicious_avoid_pct":100*(NUM_TASKS-mal_hits)/NUM_TASKS,
            "cloud_escalation_pct":100*cloud/NUM_TASKS,"avg_overhead_us":float(np.mean(overhead)),
            "avg_cpu_util_pct":100*float(np.mean(cpu)),"ml_accuracy":mean_ml["accuracy"],"ml_precision":mean_ml["precision"],
            "ml_recall":mean_ml["recall"],"ml_f1":mean_ml["f1"],"model_updates":tm.update_count,
            "window_metrics":window_metrics}

def aggregate(rows):
    output=[]
    for sc in SCENARIOS:
        for scheduler in SCHEDULERS:
            group=[r for r in rows if r["scenario"]==sc and r["scheduler"]==scheduler]
            item={"scenario":sc,"scheduler":scheduler,"runs":len(group)}
            for key in group[0]:
                if key not in {"scenario","scheduler","seed","window_metrics"}:
                    values=[r[key] for r in group]
                    item[key]=round(float(np.mean(values)),4); item[key+"_std"]=round(float(np.std(values,ddof=1)),4) if len(values)>1 else 0.0
            output.append(item)
    return output

def run_all_experiments(seeds=DEFAULT_SEEDS):
    rows=[run_scenario(sc, scheduler, seed) for seed in seeds for sc in SCENARIOS for scheduler in SCHEDULERS]
    summary=aggregate(rows)
    Path("results.json").write_text(json.dumps({"seeds":list(seeds),"per_run":rows,"summary":summary}, indent=2))
    for r in summary:
        print(f"{r['scenario']} {r['scheduler']:5s} Lat={r['avg_latency_ms']:.2f}±{r['avg_latency_ms_std']:.2f} ms | Avoid={r['malicious_avoid_pct']:.1f}% | Cloud={r['cloud_escalation_pct']:.1f}%")
    return summary

if __name__ == "__main__": run_all_experiments()
