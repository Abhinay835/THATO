"""Weight sensitivity analysis using the same event-driven, adaptive protocol."""
import matplotlib.pyplot as plt
import numpy as np
from experiment import run_scenario, DEFAULT_SEEDS

configs = [("Trust-heavy", .10, .50), ("Balanced", .30, .30), ("Latency-heavy", .50, .10)]
seeds = DEFAULT_SEEDS
scenarios = ("S5", "S6")
results = {sc: [] for sc in scenarios}
for sc in scenarios:
    for name, alpha, delta in configs:
        runs = [run_scenario(sc, "TAHTO", seed, alpha, .20, .20, delta) for seed in seeds]
        results[sc].append((name, np.mean([r["avg_latency_ms"] for r in runs]),
                            np.mean([r["malicious_avoid_pct"] for r in runs])))
fig, axes = plt.subplots(1, 2, figsize=(10, 4))
for ax, sc in zip(axes, scenarios):
    names, latency, avoid = zip(*results[sc]); x = np.arange(len(names))
    ax.bar(x-.18, latency, .36, label="Latency (ms)", color="#1976D2")
    ax2=ax.twinx(); ax2.bar(x+.18, avoid, .36, label="Avoidance (%)", color="#388E3C")
    ax.set(title=sc, xticks=x, xticklabels=names, ylabel="Latency (ms)"); ax2.set_ylabel("Avoidance (%)")
fig.tight_layout(); fig.savefig("fig6_sensitivity_analysis.png", dpi=160); plt.close(fig)
print("Saved fig6_sensitivity_analysis.png")
