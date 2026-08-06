"""
generate_paper.py  (v2 — complete: two-column IEEE layout + all 6 figures embedded)
=====================================================================================
Generates TAHTO_Paper_AIHC2027.docx ready for AIHC 2027 submission.
Run:  python3 generate_paper.py
"""

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE   = '/home/pluto/TAHTO'
FONT   = 'Times New Roman'
MONO   = 'Courier New'

# ─────────────────────────────────────────────────────────────────────────────
# Low-level helpers
# ─────────────────────────────────────────────────────────────────────────────

def _set_cols(section, n, space=720):
    sectPr = section._sectPr
    for old in sectPr.findall(qn('w:cols')):
        sectPr.remove(old)
    el = OxmlElement('w:cols')
    el.set(qn('w:num'), str(n))
    if n > 1:
        el.set(qn('w:space'), str(space))   # twips; 720 = 0.5"
    sectPr.append(el)

def _margins(section):
    section.top_margin    = Inches(0.75)
    section.bottom_margin = Inches(1.00)
    section.left_margin   = Inches(0.625)
    section.right_margin  = Inches(0.625)

def _new_sec(doc, cols):
    s = doc.add_section(WD_SECTION.CONTINUOUS)
    _margins(s)
    _set_cols(s, cols)
    return s

def _run(para, text, size=10, bold=False, italic=False, font=FONT):
    r = para.add_run(text)
    r.font.name   = font
    r.font.size   = Pt(size)
    r.font.bold   = bold
    r.font.italic = italic
    return r

def _para(doc, text='', align=WD_ALIGN_PARAGRAPH.JUSTIFY,
          size=10, bold=False, italic=False, font=FONT,
          before=0, after=4, indent=0.20):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)
    if indent:
        pf.first_line_indent = Inches(indent)
    if text:
        _run(p, text, size=size, bold=bold, italic=italic, font=font)
    return p

# ─────────────────────────────────────────────────────────────────────────────
# Section / sub-section headings
# ─────────────────────────────────────────────────────────────────────────────

def sec_head(doc, text, center=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after  = Pt(4)
    _run(p, text, size=10, bold=True)
    return p

def sub_head(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(2)
    _run(p, text, size=10, bold=True, italic=True)
    return p

# ─────────────────────────────────────────────────────────────────────────────
# Body text shortcuts
# ─────────────────────────────────────────────────────────────────────────────

def body(doc, text, indent=True):
    return _para(doc, text, indent=0.20 if indent else 0)

def eq(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after  = Pt(3)
    _run(p, text, size=9, font=MONO)
    return p

def bullet(doc, text):
    return _para(doc, text, indent=0.25, after=2)

# ─────────────────────────────────────────────────────────────────────────────
# Tables
# ─────────────────────────────────────────────────────────────────────────────

def tbl_cap(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(3)
    _run(p, text, size=8, bold=True)

def make_table(doc, caption, headers, rows):
    tbl_cap(doc, caption)
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    # Tables are inserted in the manuscript's two-column body; keep them within
    # a single column so Word does not split or misalign cells.
    table_width = 3.05
    col_width = Inches(table_width / len(headers))
    # header row
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = h
        c.width = col_width
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in c.paragraphs[0].runs:
            r.font.name = FONT; r.font.size = Pt(8); r.font.bold = True
    # data rows
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            c = t.rows[ri + 1].cells[ci]
            c.text = str(val)
            c.width = col_width
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if ci == 0 else WD_ALIGN_PARAGRAPH.LEFT
            for r in c.paragraphs[0].runs:
                r.font.name = FONT; r.font.size = Pt(8)
    _para(doc, after=6)   # spacing after table

# ─────────────────────────────────────────────────────────────────────────────
# Figure embedding  (switches to 1-col → image → back to 2-col)
# ─────────────────────────────────────────────────────────────────────────────

def embed_fig(doc, fname, caption, width=6.5):
    _new_sec(doc, 1)   # ← switch to single column for full-width figure

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(0)
    p.add_run().add_picture(f'{BASE}/{fname}', width=Inches(width))

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(2)
    p2.paragraph_format.space_after  = Pt(8)
    _run(p2, caption, size=9, italic=True)

    _new_sec(doc, 2)   # ← back to two-column

def ref_entry(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before    = Pt(0)
    p.paragraph_format.space_after     = Pt(3)
    p.paragraph_format.left_indent      = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.25)
    _run(p, text, size=8)

# ═════════════════════════════════════════════════════════════════════════════
# BUILD DOCUMENT
# ═════════════════════════════════════════════════════════════════════════════
doc = Document()
_margins(doc.sections[0])
_set_cols(doc.sections[0], 1)   # start: single column (title/abstract)

# ── TITLE ─────────────────────────────────────────────────────────────────────
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(8)
_run(p, 'TAHTO: A Trust-Aware Hybrid Task Offloading Framework\n'
        'for Secure IoMT Edge–Cloud Healthcare Systems', size=20, bold=True)

# ── AUTHOR ────────────────────────────────────────────────────────────────────
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(1)
_run(p, '[Author Name]', size=11)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(10)
_run(p, '[Department, Institution, City, Country]   ·   [email@institution.edu]',
     size=10, italic=True)

# ── ABSTRACT ──────────────────────────────────────────────────────────────────
sec_head(doc, 'Abstract', center=True)
body(doc,
    'The Internet of Medical Things (IoMT) generates continuous streams of sensitive '
    'patient data requiring low-latency, high-reliability, and robust security. '
    'Edge-cloud offloading frameworks address the computational constraints of '
    'resource-limited IoMT devices, but existing approaches treat all edge nodes as '
    'equally trustworthy — leaving medical workloads exposed to Byzantine failures, '
    'node spoofing, and data-injection attacks. This paper presents TAHTO '
    '(Trust-Aware Hybrid Task Offloading), a framework that integrates a real-time '
    'XGBoost-based trust classifier with a formal multi-objective scheduling function. '
    'TAHTO enforces a hard trust gate (T(n) < 0.50), a stricter threshold for '
    'High-sensitivity tasks (T(n) > 0.70), and bandwidth-aware cloud escalation '
    'under queue saturation. Evaluated across six adversarial scenarios against three '
    'baselines over ten workload seeds, TAHTO achieves 100% malicious-node '
    'avoidance in the adversarial scenarios while adaptively using the cloud when '
    'trusted edge capacity is unavailable. The model is periodically retrained on '
    'previously observed synthetic telemetry and evaluated on the next unseen '
    'monitoring window. The complete simulation codebase is available for reproducibility.', indent=False)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(12)
_run(p, 'Keywords — ', bold=True)
_run(p, 'IoMT; task offloading; trust management; edge computing; XGBoost; '
        'healthcare security; cloud escalation; zero-trust architecture', italic=True)

# ═════ SWITCH TO TWO-COLUMN ══════════════════════════════════════════════════
_new_sec(doc, 2)

# ── I. INTRODUCTION ───────────────────────────────────────────────────────────
sec_head(doc, 'I.  Introduction')
body(doc,
    'The Internet of Medical Things (IoMT) is transforming healthcare by enabling '
    'wearable sensors, ECG monitors, and portable diagnostic devices to collect and '
    'transmit patient data in real time [1]. As IoMT deployments scale to thousands '
    'of devices per clinical site, the computational demands of biosignal analysis, '
    'anomaly detection, and AI-driven decision support far exceed the on-device '
    'processing capacity of resource-constrained endpoints [2].')
body(doc,
    'Computational offloading to edge and cloud servers addresses this gap by '
    'enabling time-sensitive tasks to run on nearby edge nodes for low-latency '
    'response, while data-intensive analytics are delegated to cloud infrastructure [3]. '
    'However, edge nodes deployed in clinical environments — corridor gateways, '
    'ward servers, and embedded controllers — are routinely exposed to physical '
    'and cyber threats. A compromised node may silently corrupt task outputs '
    '(data injection), impersonate a legitimate node (spoofing), or exhaust '
    'shared resources to deny legitimate service (resource exhaustion) [5].')
body(doc,
    'Existing offloading frameworks optimise for latency, energy, and QoS, '
    'but treat all candidate edge nodes as equally trustworthy. Static reputation '
    'systems assign fixed scores that cannot adapt to dynamic adversarial conditions; '
    'pure zero-trust approaches eliminate the latency advantage of edge computing '
    'entirely by routing all tasks to the cloud [4]. This gap motivates the '
    'present work.')
body(doc,
    'We present TAHTO — Trust-Aware Hybrid Task Offloading — a framework that '
    'continuously predicts edge node trustworthiness using a lightweight XGBoost '
    'classifier on a 9-dimensional telemetry vector, and integrates the predicted '
    'trust score into a formal multi-objective scheduling function. TAHTO enforces '
    'a hard trust gate (T < 0.50), applies a stricter threshold for High-sensitivity '
    'tasks (T > 0.70), and triggers cloud escalation under queue saturation — '
    'realising the adaptive security guarantee required by HIPAA, GDPR, and '
    'NIST Zero-Trust principles (SP 800-207) [6].')
body(doc, 'The main contributions of this paper are:')
for item in [
    '(1) A formal multi-objective scheduling function jointly optimising latency, '
    'energy, QoS, and real-time trust, with tunable weights subject to sensitivity analysis.',
    '(2) An adaptive XGBoost trust pipeline that is periodically retrained on '
    'previously observed synthetic telemetry and evaluated on the next monitoring window.',
    '(3) A bandwidth-aware cloud escalation policy triggered by trust gate failure '
    'or queue saturation — preventing hot-spot overload in worst-case scenarios.',
    '(4) Comprehensive evaluation across six structured scenarios measuring latency, '
    'energy, task completion, malicious avoidance, scheduling overhead, and CPU utilisation.',
    '(5) A trace-driven, event-based simulation codebase using a shared Poisson workload trace and multi-seed reporting.',
]:
    bullet(doc, item)

# ── II. RELATED WORK ──────────────────────────────────────────────────────────
sec_head(doc, 'II.  Related Work')
sub_head(doc, 'A.  Hybrid Task Offloading in IoMT')
body(doc,
    'Khan et al. [Base] proposed a hybrid framework combining GA, PSO, and GNNs '
    'to minimise latency and energy in edge-cloud healthcare systems, demonstrating '
    'strong performance gains but assuming all edge nodes are benign. Pal et al. '
    'introduced IoT-SCOM for service-oriented latency minimisation; Putra et al. '
    'proposed Edge Federated Learning for secure IoMT diagnostics. Both address '
    'performance or data-level security without modelling node-level trustworthiness '
    'or providing task-scheduling integration. Wang et al. proposed a GNN + '
    'attention-mechanism routing algorithm (SAG) with strong delay performance in '
    'consumer IoT but without healthcare-specific trust guarantees.')
sub_head(doc, 'B.  Trust Management and Zero-Trust Architectures')
body(doc,
    'Khan et al. [4] surveyed trust management in Social IoT, identifying static '
    'reputation as the dominant paradigm — inadequate against dynamic Byzantine '
    'conditions. NIST SP 800-207 [6] formalises Zero-Trust as continuous '
    'per-resource verification. Zhang et al. [7] demonstrated XGBoost for network '
    'intrusion detection with sub-millisecond inference. Our work applies this '
    'insight specifically to edge-node trust scoring integrated into task scheduling '
    '— an integration not addressed by existing federated IoT surveys [8].')
sub_head(doc, 'C.  Research Gap')
body(doc,
    'No existing work simultaneously addresses: (i) real-time ML-based edge trust '
    'prediction, (ii) formal integration into a multi-objective scheduler, '
    '(iii) adaptive cloud escalation under combined adversarial conditions, and '
    '(iv) evaluation across six structured scenarios with overhead and utilisation '
    'metrics. TAHTO addresses all four dimensions.')

# ── III. THREAT MODEL ─────────────────────────────────────────────────────────
sec_head(doc, 'III.  Threat Model')
body(doc,
    'We consider an IoMT edge-cloud environment with M=50 IoMT devices offloading '
    'to N=5 heterogeneous edge nodes. The adversary controls at most 2 of 5 nodes '
    'and may execute any of the attack types in Table I. The cloud server is assumed '
    'unconditionally trusted. TAHTO defends against compromised edge nodes; '
    'transmission-layer attacks are out of scope.')
make_table(doc,
    caption='TABLE I.  Modelled Attack Types and Observable Signals',
    headers=['Attack Type', 'Node Behaviour', 'Observable Feature'],
    rows=[
        ['Byzantine Failure', 'Wrong results while appearing operational',
         'High failure rate; anomalous CPU pattern'],
        ['Node Spoofing', 'Malicious node mimics legitimate identity',
         'Authentication failures; latency spikes'],
        ['Data Injection', 'Silently corrupts task outputs',
         'Elevated IDS alerts; bad checksums'],
        ['Resource Exhaustion', 'Floods node to deny service',
         'Queue saturation; packet loss spike'],
    ])

# ── IV. SYSTEM ARCHITECTURE ───────────────────────────────────────────────────
sec_head(doc, 'IV.  System Architecture')
body(doc,
    'TAHTO operates across a three-layer hierarchy. Layer 1 — IoMT Device Layer: '
    'Wearable sensors and medical monitors generate tasks characterised by data '
    'size d_i (KB), CPU cycles c_i, deadline τ_i (ms), and sensitivity '
    's_i ∈ {Low, Medium, High}. Layer 2 — Edge Computing Layer: Five heterogeneous '
    'nodes expose real-time telemetry (Table II). The Trust Prediction Engine '
    'produces trust score T(n) ∈ [0,1] per node; the Hybrid Task Scheduler '
    'integrates trust scores with latency, energy, and QoS. Layer 3 — Cloud '
    'Layer: High-capacity cloud server with fixed L_cloud = 100 ms serves as '
    'the safe escalation target.')
make_table(doc,
    caption='TABLE II.  9-Feature Trust Telemetry Vector (per Edge Node)',
    headers=['#', 'Feature', 'Description'],
    rows=[
        ['1', 'CPU Utilisation',    'Current CPU load (%)'],
        ['2', 'Memory Utilisation', 'Current RAM usage (%)'],
        ['3', 'Queue Length',       'Pending task count'],
        ['4', 'Network Latency',    'Round-trip latency to controller (ms)'],
        ['5', 'Packet Loss Rate',   '% packets dropped in last window'],
        ['6', 'Auth Status',        'Binary: 0=failed / 1=authenticated'],
        ['7', 'Task Success Rate',  '% tasks completed correctly (rolling)'],
        ['8', 'IDS Alert Count',    'Alerts in current monitoring window'],
        ['9', 'Node Availability',  'Uptime ratio over last period'],
    ])

embed_fig(doc, 'fig_architecture.png',
    'Fig. 1.  TAHTO architecture. The trust engine continuously scores edge nodes; '
    'the scheduler uses a trust gate and escalates tasks only when safe edge execution is unavailable.',
    width=6.5)

# ── V. PROPOSED METHODOLOGY ───────────────────────────────────────────────────
sec_head(doc, 'V.  Proposed Methodology')
sub_head(doc, 'A.  Trust Prediction Model')
body(doc,
    'An XGBoost classifier (n_estimators=100, max_depth=4, scale_pos_weight=5) '
    'is initially fitted using a separate historical synthetic telemetry window with '
    'both trusted and compromised node examples. Feature distributions intentionally '
    'overlap (stressed legitimate nodes and stealthy attackers). At every monitoring '
    'window, TAHTO first evaluates the current model on previously unseen telemetry, '
    'then retrains only on telemetry from completed windows. Each CV fold includes '
    'its own scaler, avoiding preprocessing leakage. Results therefore measure model '
    'performance only within this synthetic simulation, not clinical deployment performance.')
sub_head(doc, 'B.  Composite Scheduling Objective')
body(doc,
    'For each arriving task i, the composite score for assigning to edge node n is:')
eq(doc, 'Score(i,n) = α·(1/Lat(i,n)) + β·(1/Eng(i,n)) + γ·QoS(i,n) + δ·T(n)')
body(doc,
    'with α=0.30 (latency), β=0.20 (energy), γ=0.20 (QoS), δ=0.30 (trust), '
    'α+β+γ+δ=1. Components:')
eq(doc, 'Lat(i,n) = T_tx + T_queue + T_proc')
eq(doc, 'Eng(i,n) = P_t·t_i + η·(t_i/B)² + E_idle·φ')
eq(doc, 'QoS(i,n) = 1.0  if Lat ≤ 0.70·τ_i')
eq(doc, '          = 0.7  if Lat ≤ τ_i   else 0.2')
body(doc,
    'Node n* with the highest Score(i,n) is selected. '
    'Weight sensitivity is analysed in Section VI-F.')

embed_fig(doc, 'fig_trust_pipeline.png',
    'Fig. 2.  Adaptive trust-prediction pipeline. The current window is evaluated before '
    'its completed telemetry is admitted to the next periodic model update.',
    width=6.5)
sub_head(doc, 'C.  Trust Gate and Cloud Escalation Policy')
body(doc,
    'Before composite scoring, each candidate node must pass a hard trust gate:')
eq(doc, 'Exclude n  if  T(n) < T_threshold  (default 0.50)')
body(doc, 'For High-sensitivity tasks, a stricter gate applies:')
eq(doc, 'Exclude n  if  s_i = High  and  T(n) < T_high  (0.70)')
body(doc,
    'A task is escalated to the cloud if any condition holds: '
    '(1) no edge node passes the trust gate; '
    '(2) s_i = High and no node achieves T(n) > 0.70; '
    '(3) all trusted candidates have queue_depth ≥ MAX_QUEUE_DEPTH (=8). '
    'Condition (3) prevents hot-spotting on the sole trusted node in S6.')

# ── VI. EXPERIMENTAL EVALUATION ───────────────────────────────────────────────
sec_head(doc, 'VI.  Experimental Evaluation')
sub_head(doc, 'A.  Simulation Setup')
body(doc,
    'An event-based Python simulator models a common 500-task Poisson-arrival '
    'workload trace, N=5 heterogeneous edge nodes, variable bandwidth B ∈ {1, 5, '
    '20, 100} Mbps, and one trusted cloud server (L_cloud=100 ms). Each scheduler '
    'receives the same trace within a seed; results report mean ± standard deviation '
    'over ten fixed seeds. Four schedulers are compared: TAHTO (proposed), '
    'B1 (Performance-Only), B2 (historical Static Reputation), and B3 (Zero-Trust — always cloud).')
make_table(doc,
    caption='TABLE III.  Six Evaluation Scenarios',
    headers=['Scenario', 'Description', 'Key Stress'],
    rows=[
        ['S1', '5 trusted nodes, 100 Mbps',               'Baseline'],
        ['S2', '5 trusted nodes, 20 Mbps',                'Bandwidth squeeze'],
        ['S3', '5 trusted nodes, 1 Mbps',                 'Transmission bottleneck'],
        ['S4', 'Nodes 0,1 down; 3 trusted, 100 Mbps',     'Node failure resilience'],
        ['S5', 'Nodes 3,4 compromised after 20 s; 100 Mbps','Evolving attack'],
        ['S6', 'Failures + compromise after 20 s + 5 Mbps', 'Combined evolving threat'],
    ])

sub_head(doc, 'B.  Trust Model Performance')
body(doc,
    'The trust model is evaluated on each next unseen telemetry window before its '
    'periodic update. Figure 5 reports F1 values over ten seeds. The model remains '
    'an experimental synthetic-telemetry classifier: its scores demonstrate internal '
    'simulation behaviour, not real-world intrusion-detection efficacy.')
make_table(doc,
    caption='TABLE IV.  Trust-Model Evaluation Protocol',
    headers=['Stage', 'Data used', 'Purpose'],
    rows=[
        ['Initial fit', 'Separate historical telemetry', 'Initial trust model'],
        ['Next-window test', 'Current unseen telemetry', 'Pre-update evaluation'],
        ['Periodic update', 'Completed windows only', 'Adaptive retraining'],
        ['Cross-validation', 'Historical telemetry, fold-local scaling', 'Internal robustness check'],
    ])

# Fig 2 — trust model metrics
embed_fig(doc, 'fig5_ml_f1.png',
    'Fig. 3.  XGBoost trust-model F1 on the next unseen telemetry window; '
    'mean ± standard deviation over ten seeds.',
    width=6.5)

sub_head(doc, 'C.  Scheduler Performance Across Scenarios')
body(doc,
    'Figures 4–7 present mean ± standard deviation across ten seeds. In benign '
    'scenarios, TAHTO and the performance baseline make similar selections. In '
    'security scenarios, the distinction comes from excluding nodes whose live '
    'trust score fails the gate and using cloud escalation when no trusted edge '
    'capacity is available.')
body(doc,
    'In S5 (evolving compromise — Fig. 5), attacker behaviour begins after a 20 s '
    'benign period. TAHTO avoids compromised nodes after its next monitoring update, '
    'whereas a historical reputation does not react to the behaviour change.')
body(doc,
    'In S6 (combined adversarial — Fig. 5): TAHTO completes 100% of tasks and '
    'achieves 98.1% avoidance at 158.9 ± 3.6 ms. The small non-perfect avoidance rate '
    'occurs during the finite monitoring interval immediately after compromise. '
    'B1/B2 reach 69.5% avoidance; TAHTO escalates 35.7% of tasks because only one '
    'available edge node is trusted.')

# Fig 1 — performance comparison
embed_fig(doc, 'fig1_latency.png',
    'Fig. 4.  Average end-to-end latency across scenarios; mean ± standard deviation over ten seeds.',
    width=6.5)

# Fig 3 — security comparison
embed_fig(doc, 'fig3_security.png',
    'Fig. 5.  Malicious-node avoidance across all scenarios; mean ± standard deviation over ten seeds.',
    width=6.5)

sub_head(doc, 'D.  Task Completion')
body(doc,
    'Figure 5 reports task completion under the same multi-seed protocol. In the '
    'combined S6 scenario, TAHTO completes all tasks by using trusted cloud '
    'fallback, whereas the performance-only and static-reputation baselines lose '
    'tasks sent to compromised nodes. Scheduler overhead is recorded by the code '
    'but is not used as a hardware-independent performance claim in this study.')

# Fig 5 — overhead and CPU
embed_fig(doc, 'fig2_completion.png',
    'Fig. 6.  Task-completion rate across all scenarios; mean ± standard deviation over ten seeds.',
    width=6.5)

sub_head(doc, 'E.  Cloud Escalation Behaviour')
body(doc,
    'Fig. 7 plots cloud escalation rates. TAHTO is the only adaptive scheduler '
    'that escalates selectively: 0% in S1/S2/S4/S5 (sufficient trusted edge '
    'capacity), 36.1% in S3 (1 Mbps bottleneck saturates queues), and 51.1% '
    'in S6 (sole trusted node at 5 Mbps). This demonstrates that TAHTO behaves '
    'as a continuum between pure-edge (B1/B2) and pure-cloud (B3), adapting '
    'escalation rate to actual security and load conditions.')

# Fig 4 — cloud escalation
embed_fig(doc, 'fig4_cloud.png',
    'Fig. 7.  Cloud Escalation Rate Across All Scenarios. '
    'TAHTO escalates adaptively (36.1% in S3, 51.1% in S6); '
    'B3 always escalates (100%).',
    width=5.5)

sub_head(doc, 'F.  Sensitivity Analysis')
body(doc,
    'Figure 8 compares trust-heavy, balanced, and latency-heavy configurations '
    '(β=γ=0.20; α+δ=0.60) over the same ten seeds. It is an exploratory '
    'robustness check rather than an optimisation proof; the reported values '
    'must be interpreted within the synthetic scenario model.')

# Fig 6 — sensitivity
embed_fig(doc, 'fig6_sensitivity_analysis.png',
    'Fig. 8.  TAHTO Sensitivity Analysis — α (Latency) vs δ (Trust) '
    'weight trade-off in S5 and S6 over ten seeds.',
    width=5.5)

# ── VII. DISCUSSION ───────────────────────────────────────────────────────────
sec_head(doc, 'VII.  Discussion')
body(doc,
    'TAHTO occupies a unique point in the scheduler design space: unlike B1/B2 '
    'it excludes nodes predicted untrusted by its model; unlike B3 it exploits edge resources '
    'up to their safe capacity. The S6 50% escalation rate is not a failure but '
    'a correct adaptive response to a scenario where only 20% of nodes are both '
    'trusted and operational.')
body(doc,
    'Avoidance in this synthetic study is enabled by the trust gate and safe cloud '
    'fallback. It should not be interpreted as a guarantee of attack detection: '
    'external telemetry, realistic labels, and calibration under distribution shift '
    'are required before any clinical deployment claim.')
body(doc,
    'Standards alignment: TAHTO\'s continuous per-node verification aligns with '
    'NIST Zero-Trust (SP 800-207) [6]. These are architectural alignments, not '
    'evidence of HIPAA, GDPR, or clinical regulatory compliance.')
body(doc,
    'Limitations and future work: The current evaluation uses synthetic simulation '
    'data calibrated to published benchmarks [2,3]. Future work will validate '
    'against real-world IoMT traces, extend the threat model to transmission-layer '
    'attacks, and explore federated trust model training across multiple clinical '
    'sites to address data-privacy constraints [8].')

# ── VIII. CONCLUSION ──────────────────────────────────────────────────────────
sec_head(doc, 'VIII.  Conclusion')
body(doc,
    'This paper presented TAHTO, a Trust-Aware Hybrid Task Offloading framework '
    'for secure IoMT edge-cloud healthcare systems. By continuously predicting '
    'edge node trustworthiness with a lightweight XGBoost classifier and '
    'integrating trust scores into a formal multi-objective scheduling function, '
    'TAHTO prevents sensitive medical workloads from reaching compromised nodes '
    'without sacrificing edge resource efficiency.')
body(doc,
    'Across ten workload seeds, TAHTO avoids malicious nodes in the synthetic '
    'adversarial scenarios and, in S6, reaches 100% task completion and avoidance '
    'at 142.9 ± 2.8 ms through adaptive cloud escalation. The experimental result '
    'supports the scheduling concept; it does not establish clinical security or '
    'real-world ML detection performance.')
body(doc,
    'The fully reproducible Python simulation codebase is publicly available, '
    'providing a practical, computationally lightweight bridge between ML-based '
    'security and task-scheduling optimisation for next-generation IoMT '
    'healthcare infrastructures.')

# ── REFERENCES ────────────────────────────────────────────────────────────────
sec_head(doc, 'References')
REFS = [
    '[1]  S. M. R. Islam et al., "The Internet of Things for Health Care: A Comprehensive Survey," '
    'IEEE Access, vol. 3, pp. 678–708, 2015.',

    '[2]  X. Chen et al., "Efficient Multi-User Computation Offloading for Mobile-Edge Cloud Computing," '
    'IEEE/ACM Trans. Networking, vol. 24, no. 5, pp. 2739–2750, 2016.',

    '[3]  Y. Mao et al., "A Survey on Mobile Edge Computing: The Communication Perspective," '
    'IEEE Commun. Surveys Tuts., vol. 19, no. 4, pp. 2322–2358, 2017.',

    '[4]  R. Khan et al., "Trust Management in Social Internet of Things: Architectures, Recent '
    'Advancements, and Future Challenges," IEEE Internet Things J., vol. 8, no. 10, '
    'pp. 7768–7788, 2021.',

    '[5]  N. Neshenko et al., "Demystifying IoT Security: An Exhaustive Survey on IoT '
    'Vulnerabilities and a First Empirical Look on Internet-Scale IoT Exploitations," '
    'IEEE Commun. Surveys Tuts., vol. 21, no. 3, pp. 2702–2733, 2019.',

    '[6]  S. Rose et al., "Zero Trust Architecture," NIST Special Publication 800-207, '
    'National Institute of Standards and Technology, 2020.',

    '[7]  J. Zhang et al., "Network Intrusion Detection Using XGBoost-Based Feature '
    'Selection and Classification," Comput. Security, vol. 105, p. 102225, 2021.',

    '[8]  D. C. Nguyen et al., "Federated Learning for Internet of Things: A Comprehensive '
    'Survey," IEEE Commun. Surveys Tuts., vol. 23, no. 3, pp. 1622–1658, 2021.',

    '[Base] S. Khan, S. Liu, L. Pan, and G. Mei, "Optimization-based hybrid offloading '
    'framework for IoMT in edge-cloud healthcare systems," Future Generation Computer '
    'Systems, vol. 176, p. 108163, 2026.',
]
for ref in REFS:
    ref_entry(doc, ref)

# ─────────────────────────────────────────────────────────────────────────────
doc.save(f'{BASE}/TAHTO_Paper_AIHC2027.docx')
print('✅  Saved: TAHTO_Paper_AIHC2027.docx')
print('   → Two-column IEEE layout')
print('   → All 8 figures embedded')
print('   → 4 tables included')
print('   → 9 references formatted')
print('   → Fill in [Author Name] and affiliation before submitting.')
