"""Create publication-ready TAHTO architecture and adaptive-trust pipeline diagrams."""
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

BLUE, GREEN, ORANGE, RED, GREY = "#1976D2", "#2E7D32", "#EF6C00", "#C62828", "#546E7A"

def box(ax, x, y, w, h, text, color, fontsize=10):
    patch = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.025",
                           facecolor=color, edgecolor="#263238", linewidth=1.2, alpha=.94)
    ax.add_patch(patch)
    ax.text(x+w/2, y+h/2, text, ha="center", va="center", color="white",
            fontsize=fontsize, fontweight="bold", wrap=True)

def arrow(ax, start, end, label="", color=GREY):
    ax.add_patch(FancyArrowPatch(start, end, arrowstyle="-|>", mutation_scale=15,
                                 linewidth=1.4, color=color))
    if label:
        ax.text((start[0]+end[0])/2, (start[1]+end[1])/2+.025, label, ha="center",
                va="bottom", fontsize=8, color=color)

def architecture():
    fig, ax = plt.subplots(figsize=(11, 6.5)); ax.set(xlim=(0, 1), ylim=(0, 1)); ax.axis("off")
    ax.text(.5, .96, "TAHTO Architecture: Trust-Aware IoMT Edge–Cloud Offloading", ha="center",
            fontsize=15, fontweight="bold")
    box(ax, .07, .10, .22, .20, "IoMT Device Layer\nWearables • ECG • Sensors\nTasks: size, CPU, deadline, sensitivity", BLUE, 9)
    box(ax, .39, .62, .24, .17, "Trust Prediction Engine\n9 telemetry features\nXGBoost trust score T(n)", ORANGE, 10)
    box(ax, .39, .32, .24, .17, "TAHTO Scheduler\nTrust gate + normalized\nlatency / energy / QoS score", GREEN, 10)
    box(ax, .39, .05, .24, .13, "Edge Nodes\nTrusted nodes execute tasks\nLow-trust nodes excluded", GREY, 10)
    box(ax, .73, .30, .21, .20, "Trusted Cloud\nFallback for no safe edge\nor queue saturation", BLUE, 10)
    arrow(ax, (.29,.20), (.39,.40), "task")
    arrow(ax, (.51,.18), (.51,.32), "safe assignment")
    arrow(ax, (.51,.62), (.51,.49), "live telemetry")
    arrow(ax, (.63,.405), (.73,.405), "escalate when needed")
    arrow(ax, (.51,.49), (.51,.62), "trust decision", ORANGE)
    ax.text(.51,.86,"Monitoring window", ha="center", fontsize=9, color=ORANGE, fontweight="bold")
    fig.tight_layout(); fig.savefig("fig_architecture.png", dpi=200, bbox_inches="tight"); plt.close(fig)

def pipeline():
    fig, ax = plt.subplots(figsize=(12, 4.6)); ax.set(xlim=(0,1), ylim=(0,1)); ax.axis("off")
    ax.text(.5,.94,"Adaptive Trust-Prediction and Scheduling Pipeline",ha="center",fontsize=15,fontweight="bold")
    items=[(.03,"Historical\ntelemetry",BLUE),(.21,"Initial XGBoost\nfit",ORANGE),(.39,"Next unseen\nmonitoring window",GREY),(.58,"Trust score +\nhard gate",GREEN),(.76,"Edge / cloud\noffloading",BLUE)]
    for x, label, color in items: box(ax,x,.43,.16,.22,label,color,10)
    for (x,_,_),(nx,_,_) in zip(items,items[1:]): arrow(ax,(x+.16,.54),(nx,.54))
    arrow(ax,(.47,.43),(.29,.18),"labelled completed window",ORANGE)
    arrow(ax,(.29,.18),(.29,.43),"periodic retraining",ORANGE)
    ax.text(.29,.08,"Only past windows update the model; current-window metrics are recorded first.",ha="center",fontsize=9,color=ORANGE)
    fig.tight_layout(); fig.savefig("fig_trust_pipeline.png", dpi=200, bbox_inches="tight"); plt.close(fig)

if __name__ == "__main__":
    architecture(); pipeline(); print("Saved fig_architecture.png and fig_trust_pipeline.png")
