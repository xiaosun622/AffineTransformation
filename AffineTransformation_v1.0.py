import plotly.graph_objects as go
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# -------------------------------
# File selection
# -------------------------------

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(
    title="Select Excel file", filetypes=[("Excel files", "*.xlsx *.xls")]
)

df = pd.read_excel(file_path)

# -------------------------------
# COLUMN mapping
# -------------------------------

point_id = df.iloc[:, 0].astype(str)

# CRITICAL FIX: normalise labels
point_id = point_id.str.strip().str.upper()  # remove spaces  # enforce consistent case

x1, y1 = df.iloc[:, 1], df.iloc[:, 2]
x2, y2 = df.iloc[:, 4], df.iloc[:, 5]
x3, y3 = df.iloc[:, 7], df.iloc[:, 8]

# -------------------------------
# Label classification (FIXED)
# -------------------------------

dark_set = {"R1", "R2", "R3"}


def is_dark(label):
    return label in dark_set


# -------------------------------
# Colour rules
# -------------------------------


def colour(system, dark):
    if system == 1:
        return "darkred" if dark else "lightcoral"
    if system == 2:
        return "darkblue" if dark else "lightblue"
    if system == 3:
        return "darkgreen" if dark else "lightgreen"


colors1 = [colour(1, is_dark(i)) for i in point_id]
colors2 = [colour(2, is_dark(i)) for i in point_id]
colors3 = [colour(3, is_dark(i)) for i in point_id]

# -------------------------------
# Z levels
# -------------------------------

z1, z2, z3 = 0, 10, 20

# -------------------------------
# Plot
# -------------------------------

fig = go.Figure()

fig.add_trace(
    go.Scatter3d(
        x=x1,
        y=y1,
        z=[z1] * len(df),
        mode="markers+text",
        marker=dict(size=5, color=colors1),
        text=point_id,
        textfont=dict(
            size=14,
            # family="Arial Black"
        ),
        textposition="top center",
        name="System 1",
    )
)

fig.add_trace(
    go.Scatter3d(
        x=x2,
        y=y2,
        z=[z2] * len(df),
        mode="markers+text",
        marker=dict(size=5, color=colors2),
        text=point_id,
        textfont=dict(
            size=14,
            # family="Arial Black"
        ),
        textposition="top center",
        name="System 2",
    )
)

fig.add_trace(
    go.Scatter3d(
        x=x3,
        y=y3,
        z=[z3] * len(df),
        mode="markers+text",
        marker=dict(size=5, color=colors3),
        text=point_id,
        textfont=dict(
            size=14,
            # family="Arial Black"
        ),
        textposition="top center",
        name="System 3",
    )
)

# -------------------------------
# Background planes
# -------------------------------

xmin = np.nanmin([x1.min(), x2.min(), x3.min()])
xmax = np.nanmax([x1.max(), x2.max(), x3.max()])
ymin = np.nanmin([y1.min(), y2.min(), y3.min()])
ymax = np.nanmax([y1.max(), y2.max(), y3.max()])

X, Y = np.meshgrid(np.linspace(xmin, xmax, 15), np.linspace(ymin, ymax, 15))

for z in [z1, z2, z3]:
    fig.add_trace(
        go.Surface(x=X, y=Y, z=np.full_like(X, z), opacity=0.18, showscale=False)
    )

# -------------------------------
# Layout
# -------------------------------

fig.update_layout(
    title="3D Layered Coordinate System Comparison",
    scene=dict(
        xaxis=dict(
            title=dict(text="X", font=dict(size=22)),
            showticklabels=False,
            showgrid=False,
            zeroline=False,
        ),
        yaxis=dict(
            title=dict(text="Y", font=dict(size=22)),
            showticklabels=False,
            showgrid=False,
            zeroline=False,
        ),
        zaxis=dict(
            title=dict(text="Z", font=dict(size=22)),
            showticklabels=False,
            showgrid=False,
            zeroline=False,
        ),
    ),
)

fig.show()
