"""Event-driven IoMT edge--cloud simulator used by the TAHTO study."""
from __future__ import annotations

from dataclasses import dataclass
import numpy as np

CLOUD_LATENCY = 100.0
TRUST_THRESHOLD = 0.50
HIGH_SENS_THRESHOLD = 0.70
MAX_QUEUE_DEPTH = 8


@dataclass(frozen=True)
class Task:
    task_id: int
    arrival_ms: float
    size_kb: float
    cpu_cycles: float
    deadline_ms: float
    sensitivity: str


def generate_workload(num_tasks=500, rate_per_second=8.0, seed=42):
    """Generate a common Poisson-arrival workload trace for every scheduler."""
    rng = np.random.default_rng(seed)
    arrivals = np.cumsum(rng.exponential(1000.0 / rate_per_second, num_tasks))
    tasks = []
    for i, arrival in enumerate(arrivals):
        sensitivity = rng.choice(["Low", "Medium", "High"], p=[.45, .35, .20])
        tasks.append(Task(i, float(arrival), float(rng.uniform(10, 200)),
                          float(rng.uniform(1e6, 1e8)), float(rng.uniform(50, 200)),
                          str(sensitivity)))
    return tasks


class EdgeNode:
    def __init__(self, node_id, cpu_freq_ghz, bandwidth_mbps, malicious=False,
                 availability=1.0, attack_start_ms=0.0, seed=42):
        self.node_id, self.cpu_freq, self.bandwidth = node_id, cpu_freq_ghz, bandwidth_mbps
        self.malicious, self.availability = malicious, availability
        self.attack_start_ms = attack_start_ms
        self.rng = np.random.default_rng(seed + node_id * 101)
        self.available_at = 0.0
        self.trust_score, self.static_reputation = .9, .9
        self.cpu_util = self.mem_util = self.packet_loss = 0.0
        self.queue_length = self.ids_alert_count = 0
        self.net_latency, self.auth_status = 5.0, 1
        self.task_success_rate = self.node_availability = 1.0

    def is_compromised(self, now):
        """Compromise becomes active after a scenario-defined change point."""
        return self.malicious and now >= self.attack_start_ms

    def queue_depth(self, now):
        # Approximate in-flight work from service backlog.
        return max(0, int(np.ceil((self.available_at - now) / 8.0)))

    def update_telemetry(self, now):
        q = min(self.queue_depth(now), MAX_QUEUE_DEPTH + 4)
        if self.availability < .1:
            self.cpu_util, self.mem_util, self.queue_length = .0, .0, q
            self.net_latency, self.packet_loss, self.auth_status = 100., 1., 0
            self.task_success_rate, self.ids_alert_count, self.node_availability = .0, 12, .0
        elif self.is_compromised(now):
            # Overlaps stressed legitimate-node telemetry; attacks are not label-leaking.
            self.cpu_util = float(self.rng.uniform(.35, .88))
            self.mem_util = float(self.rng.uniform(.30, .82))
            self.queue_length = int(q + self.rng.integers(0, 5))
            self.net_latency = float(self.rng.uniform(5, 42))
            self.packet_loss = float(self.rng.uniform(.06, .30))
            self.auth_status = int(self.rng.random() > .60)
            self.task_success_rate = float(self.rng.uniform(.38, .78))
            self.ids_alert_count = int(self.rng.integers(3, 10))
            self.node_availability = float(self.rng.uniform(.50, .85))
        else:
            load = min(q / MAX_QUEUE_DEPTH, 1.0)
            self.cpu_util = float(min(.98, self.rng.uniform(.20, .63) + .30 * load))
            self.mem_util = float(self.rng.uniform(.20, .65))
            self.queue_length = int(q + self.rng.integers(0, 3))
            self.net_latency = float(self.rng.uniform(1, 25))
            self.packet_loss = float(self.rng.uniform(0, .10 if load > .7 else .04))
            self.auth_status = int(self.rng.random() > .03)
            self.task_success_rate = float(self.rng.uniform(.72 if load > .7 else .86, 1.0))
            self.ids_alert_count = int(self.rng.integers(0, 4 if load > .7 else 2))
            self.node_availability = float(self.rng.uniform(.84, 1.0))

    def features(self):
        return [self.cpu_util, self.mem_util, self.queue_length, self.net_latency,
                self.packet_loss, self.auth_status, self.task_success_rate,
                self.ids_alert_count, self.node_availability]

    def label(self, now):
        return int(not self.is_compromised(now) and self.availability >= .1)

    def estimate(self, task, now):
        tx = task.size_kb * 8 / (self.bandwidth * 1000) * 1000
        proc = task.cpu_cycles / (self.cpu_freq * 1e9) * 1000
        start = max(now + tx, self.available_at)
        finish = start + proc
        return finish - now, finish

    def energy(self, task):
        tx_seconds = task.size_kb * 8 / (self.bandwidth * 1e6)
        return .5 * tx_seconds + .03

    def qos(self, latency, deadline):
        return 1.0 if latency <= .7 * deadline else (.7 if latency <= deadline else .2)

    def assign(self, task, now):
        latency, finish = self.estimate(task, now)
        self.available_at = finish
        return latency, self.energy(task)


def normalized_score(candidates, alpha, beta, gamma, delta):
    """Min--max normalisation makes scheduler weights interpretable."""
    lat = np.array([x[2] for x in candidates]); energy = np.array([x[3] for x in candidates])
    def benefit(values):
        span = values.max() - values.min()
        return np.ones_like(values) if span < 1e-12 else (values.max() - values) / span
    l_b, e_b = benefit(lat), benefit(energy)
    return [alpha*l_b[i] + beta*e_b[i] + gamma*x[4] + delta*x[1].trust_score
            for i, x in enumerate(candidates)]


def choose_tahto(task, nodes, now, alpha=.30, beta=.20, gamma=.20, delta=.30):
    candidates = []
    for node in nodes:
        if node.availability < .1 or node.trust_score < TRUST_THRESHOLD:
            continue
        if task.sensitivity == "High" and node.trust_score < HIGH_SENS_THRESHOLD:
            continue
        if node.queue_depth(now) >= MAX_QUEUE_DEPTH:
            continue
        latency, _ = node.estimate(task, now)
        candidates.append((0, node, latency, node.energy(task), node.qos(latency, task.deadline_ms)))
    if not candidates:
        return "cloud", CLOUD_LATENCY, .05
    scores = normalized_score(candidates, alpha, beta, gamma, delta)
    node = candidates[int(np.argmax(scores))][1]
    latency, energy = node.assign(task, now)
    return node.node_id, latency, energy


def choose_performance(task, nodes, now):
    candidates = [(n.estimate(task, now)[0], n) for n in nodes if n.availability >= .1 and n.queue_depth(now) < MAX_QUEUE_DEPTH]
    if not candidates: return "cloud", CLOUD_LATENCY, .05
    node = min(candidates, key=lambda x: x[0])[1]
    latency, energy = node.assign(task, now)
    return node.node_id, latency, energy


def choose_static_reputation(task, nodes, now):
    candidates = []
    for node in nodes:
        if node.availability < .1 or node.queue_depth(now) >= MAX_QUEUE_DEPTH: continue
        latency, _ = node.estimate(task, now)
        candidates.append((node, latency, node.energy(task), node.qos(latency, task.deadline_ms)))
    if not candidates: return "cloud", CLOUD_LATENCY, .05
    lat = np.array([x[1] for x in candidates]); e = np.array([x[2] for x in candidates])
    def b(v): return np.ones_like(v) if np.ptp(v) < 1e-12 else (v.max()-v)/np.ptp(v)
    scores = .35*b(lat) + .20*b(e) + .15*np.array([x[3] for x in candidates]) + .30*np.array([x[0].static_reputation for x in candidates])
    node = candidates[int(np.argmax(scores))][0]
    latency, energy = node.assign(task, now)
    return node.node_id, latency, energy
