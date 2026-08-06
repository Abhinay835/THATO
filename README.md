# TAHTO — Trust-Aware Hybrid Task Offloading

**Paper:** *TAHTO: A Trust-Aware Hybrid Task Offloading Framework for Secure IoMT Edge–Cloud Healthcare Systems*
**Submitted to:** IEEE International Conference on AI in Healthcare (AIHC 2027)
**Conference Record:** #71011 · Technically Co-Sponsored by IEEE Hyderabad Section

---

## Overview

TAHTO is a Python discrete-event simulation framework that evaluates a trust-aware task-offloading scheduler for Internet of Medical Things (IoMT) edge–cloud environments. It integrates an XGBoost-based real-time trust classifier with a multi-objective scheduling function to prevent sensitive medical workloads from being assigned to compromised or unreliable edge nodes.

### Key Contributions

| Contribution | Detail |
|---|---|
| Trust Prediction | Periodically retrained XGBoost classifier on a 9-feature node telemetry vector; leakage-free CV and next-window evaluation |
| Composite Scheduler | α·(1/Latency) + β·(1/Energy) + γ·QoS + δ·Trust |
| Cloud Escalation | Triggered by trust gate failure, High-sensitivity tasks, or queue saturation |
| Evaluation | Trace-driven Poisson workload, 6 scenarios × 4 baselines × 10 seeds; mean ± SD reporting |
| Sensitivity Analysis | α vs δ weight sweep across S5 and S6 |

---

## Repository Structure

```
TAHTO/
├── simulator.py            # EdgeNode model, task generator, 4 schedulers
├── trust_model.py          # XGBoost trust classifier + 10-fold CV
├── experiment.py           # 6-scenario × 4-scheduler experiment runner
├── plot_results.py         # Generates fig1–fig5
├── sensitivity_analysis.py # Weight sensitivity analysis → fig6
├── generate_diagrams.py    # Architecture and adaptive-trust pipeline figures
├── requirements.txt        # Python dependencies
│
├── results.json            # Per-run and aggregated multi-seed results
├── fig1_latency.png        # Latency (mean ± SD)
├── fig2_completion.png     # Task completion (mean ± SD)
├── fig3_security.png       # Malicious-node avoidance (mean ± SD)
├── fig4_cloud.png          # Cloud escalation (mean ± SD)
├── fig5_ml_f1.png          # Trust model's next-window F1 (mean ± SD)
└── fig6_sensitivity_analysis.png     # α vs δ trade-off (S5, S6)
```

---

## Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the full experiment (trains model + prints all results)

```bash
python3 experiment.py
```

Expected output: multi-seed aggregate results for all scenarios and schedulers. Runtime: ~20 seconds.

### 3. Generate all figures (fig1–fig5)

```bash
python3 plot_results.py
```

### 4. Run sensitivity analysis (fig6)

```bash
python3 sensitivity_analysis.py
```

### 5. Generate the architecture diagrams

```bash
python3 generate_diagrams.py
```

---

## Simulation Parameters

| Parameter | Value |
|---|---|
| IoMT Devices | M = 50 (modelled as 500 Poisson-arrival tasks) |
| Edge Nodes | N = 5 per scenario |
| Cloud Latency | 100 ms (fixed) |
| Trust Threshold | T_threshold = 0.5 (hard gate) |
| High-Sensitivity Threshold | T_high = 0.70 |
| Queue Saturation Limit | MAX_QUEUE_DEPTH = 8 tasks/node |
| Telemetry Update Interval | Δt = every 10 tasks |
| Scheduler Weights | α=0.30, β=0.20, γ=0.20, δ=0.30 |

---

## Scenarios

| Scenario | Description | Key Stress |
|---|---|---|
| S1 | Balanced — 5 trusted nodes, 100 Mbps | Baseline |
| S2 | Heavy load — 20 Mbps | Bandwidth squeeze |
| S3 | Low bandwidth — 1 Mbps | Transmission bottleneck |
| S4 | Node failures — nodes 0,1 down | Resilience |
| S5 | Evolving compromise — nodes 3,4 become Byzantine after 20 s | Security adaptation |
| S6 | Combined — failures + evolving compromise + 5 Mbps | Worst case |

---

## Baselines

| Baseline | Description |
|---|---|
| TAHTO | Trust-Aware Hybrid Task Offloading (proposed) |
| B1 (Perf-Only) | Lowest latency only — no trust awareness |
| B2 (Static Rep) | Fixed, differentiated historical reputation; cannot react after compromise |
| B3 (Zero-Trust) | All tasks unconditionally escalated to cloud |

---

## Trust Model Feature Vector

| # | Feature | Description |
|---|---|---|
| 1 | CPU Utilisation | Current CPU load (%) |
| 2 | Memory Utilisation | Current RAM usage (%) |
| 3 | Queue Length | Number of pending tasks |
| 4 | Network Latency | Round-trip latency to controller (ms) |
| 5 | Packet Loss Rate | % packets dropped in last window |
| 6 | Authentication Status | Binary: 0 = failed, 1 = authenticated |
| 7 | Task Success Rate | % tasks completed correctly (rolling window) |
| 8 | IDS Alert Count | Number of IDS alerts in current monitoring window |
| 9 | Node Availability | Uptime ratio over last monitoring period |

---

## Reproducibility

Each scheduler is evaluated against the same seeded Poisson-arrival workload trace within each scenario. Results are aggregated across ten fixed seeds and saved in `results.json`. The trust model is first fitted to a separate historical calibration window, evaluated on the next unseen monitoring window, then periodically retrained using only previously observed labelled telemetry. This is a synthetic simulation study; its ML metrics must not be interpreted as clinical or real-world intrusion-detection accuracy.

---

## Citation

If you use this simulator, please cite:

```
[Author]. TAHTO: A Trust-Aware Hybrid Task Offloading Framework for Secure
IoMT Edge–Cloud Healthcare Systems. IEEE International Conference on AI in
Healthcare (AIHC 2027), Hyderabad, India, January 2027.
```

---

## License

This code is released for academic research purposes in conjunction with the AIHC 2027 paper submission.
