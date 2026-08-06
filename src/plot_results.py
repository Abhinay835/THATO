"""Plot the multi-seed experiment summary produced by experiment.py."""
import json
from pathlib import Path
import matplotlib.pyplot as plt
import numpy as np

if not Path("results.json").exists():
    from experiment import run_all_experiments
    run_all_experiments()
result_file = json.loads(Path("results.json").read_text())
data = result_file["summary"]
seed_count = len(result_file["seeds"])
scenarios = ["S1", "S2", "S3", "S4", "S5", "S6"]
schedulers = ["TAHTO", "B1", "B2", "B3"]
labels = ["TAHTO", "Performance-only", "Static reputation", "Cloud-only"]
colors = ["#1976D2", "#D32F2F", "#F57C00", "#388E3C"]

def val(sc, scheduler, metric):
    row = next(r for r in data if r["scenario"] == sc and r["scheduler"] == scheduler)
    return row[metric], row.get(metric + "_std", 0)

def grouped(metric, ylabel, filename):
    x, width = np.arange(len(scenarios)), .19
    fig, ax = plt.subplots(figsize=(11, 5))
    for i, scheduler in enumerate(schedulers):
        values, errors = zip(*(val(sc, scheduler, metric) for sc in scenarios))
        ax.bar(x + (i - 1.5) * width, values, width, yerr=errors, capsize=3,
               label=labels[i], color=colors[i], edgecolor="black", linewidth=.4)
    ax.set(xticks=x, xticklabels=scenarios, xlabel="Scenario", ylabel=ylabel)
    ax.grid(axis="y", alpha=.25); ax.set_axisbelow(True); ax.legend(fontsize=8)
    fig.tight_layout(); fig.savefig(filename, dpi=160); plt.close(fig)

grouped("avg_latency_ms", f"Average end-to-end latency (ms), mean ± SD (n={seed_count})", "fig1_latency.png")
grouped("task_completion_pct", f"Task completion (%), mean ± SD (n={seed_count})", "fig2_completion.png")
grouped("malicious_avoid_pct", f"Malicious-node avoidance (%), mean ± SD (n={seed_count})", "fig3_security.png")
grouped("cloud_escalation_pct", f"Cloud escalation (%), mean ± SD (n={seed_count})", "fig4_cloud.png")
grouped("ml_f1", f"Trust-model F1 on next telemetry window, mean ± SD (n={seed_count})", "fig5_ml_f1.png")
print("Saved fig1_latency.png through fig5_ml_f1.png")
