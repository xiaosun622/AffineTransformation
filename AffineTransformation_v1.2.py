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

if not file_path:
    raise ValueError("No file selected")

df = pd.read_excel(file_path)

# -------------------------------
# Column mapping
# -------------------------------

point_id = df.iloc[:, 0].astype(str).str.strip()

x1, y1 = df.iloc[:, 1], df.iloc[:, 2]
x2, y2 = df.iloc[:, 4], df.iloc[:, 5]
x3, y3 = df.iloc[:, 7], df.iloc[:, 8]

# -------------------------------
# Label groups
# -------------------------------

R_set = {"R1", "R2", "R3"}

is_R = point_id.isin(R_set)
is_P = ~is_R

# -------------------------------
# Z levels
# -------------------------------

z1, z2, z3 = 0, 40, 70

# -------------------------------
# Helper for traces
# -------------------------------


def make_trace(x, y, z, mask, color, name, bold=False):
    return go.Scatter3d(
        x=x[mask],
        y=y[mask],
        z=np.full(mask.sum(), z),
        mode="markers+text",
        marker=dict(size=5, color=color),
        text=point_id[mask],
        textfont=dict(size=14, family="Arial Black" if bold else "Arial"),
        textposition="top center",
        name=name,
    )


# -------------------------------
# Build figure
# -------------------------------

fig = go.Figure()

# System 1
fig.add_trace(make_trace(x1, y1, z1, is_R, "darkred", "S1 - R", bold=True))
fig.add_trace(make_trace(x1, y1, z1, is_P, "lightcoral", "S1 - P"))

# System 2
fig.add_trace(make_trace(x2, y2, z2, is_R, "indigo", "S2 - R", bold=True))
fig.add_trace(make_trace(x2, y2, z2, is_P, "plum", "S2 - P"))

# System 3
fig.add_trace(make_trace(x3, y3, z3, is_R, "darkgreen", "S3 - R", bold=True))
fig.add_trace(make_trace(x3, y3, z3, is_P, "lightgreen", "S3 - P"))

# -------------------------------
# Background planes bounds
# -------------------------------

xmin = np.nanmin([x1.min(), x2.min(), x3.min()])
xmax = np.nanmax([x1.max(), x2.max(), x3.max()])
ymin = np.nanmin([y1.min(), y2.min(), y3.min()])
ymax = np.nanmax([y1.max(), y2.max(), y3.max()])

X, Y = np.meshgrid(np.linspace(xmin, xmax, 15), np.linspace(ymin, ymax, 15))

# -------------------------------
# Background planes
# -------------------------------

for z in [z1, z2, z3]:
    fig.add_trace(
        go.Surface(x=X, y=Y, z=np.full_like(X, z), opacity=0.18, showscale=False)
    )

# -------------------------------
# Layout (FIXED AXIS FONT HANDLING)
# -------------------------------

fig.update_layout(
    title="3D Layered Coordinate System Comparison",
    width=1100,
    height=850,
    scene=dict(
        aspectmode="cube",
        xaxis=dict(
            title=dict(text="X"),
            title_font=dict(size=26),
            showticklabels=False,
            showgrid=False,
            zeroline=False,
        ),
        yaxis=dict(
            title=dict(text="Y"),
            title_font=dict(size=26),
            showticklabels=False,
            showgrid=False,
            zeroline=False,
        ),
        zaxis=dict(
            title=dict(text=""),
            title_font=dict(size=26),
            showticklabels=False,
            showgrid=False,
            zeroline=False,
        ),
    ),
)

fig.show()
