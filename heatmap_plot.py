import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import Rectangle
import matplotlib.colors as mcolors

# Load your Excel file
df = pd.read_excel("heatmap.xlsx")

# Drop non-numeric columns (Rs ID, etc.)
numeric_df = df.drop(columns=["Rs ID"])

# Pearson correlation matrix
corr = numeric_df.corr(method="pearson")

# 3. Mask upper triangle
mask = np.triu(np.ones_like(corr, dtype=bool))

# 4. Reverse columns only (mirror horizontally)
corr_matrix = corr.iloc[:, ::-1]
mask = mask[:, ::-1]

pure_red_blue = LinearSegmentedColormap.from_list("pure_rb", ["#0000FF", "white", "#FF0000"])
# 5. Plot heatmap with pure red–blue shades
plt.figure(figsize=(8, 6))
ax = sns.heatmap(corr_matrix,mask=mask,cmap= pure_red_blue,vmin=-1,vmax=1,annot=False,linewidths=0,cbar_kws={"label": ""},square=True)
ax.add_patch(Rectangle(
    (0, 0),                       # bottom left corner
    corr_matrix.shape[1],         # width
    corr_matrix.shape[0],         # height
    fill=False,                   # no fill, just border
    edgecolor='black',            # border color
    lw=2                          # border thickness
))
# After plotting the heatmap
cbar = ax.collections[0].colorbar

cbar.outline.set_linewidth(1)     # thickness of colorbar border
cbar.outline.set_edgecolor('black')  # border color
cbar.set_ticks([1, 0.6, 0.2, -0.2, -0.6, -1])
cbar.set_ticklabels([1, 0.6, 0.2, -0.2, -0.6, -1])
plt.xticks(rotation=60, ha="right")
plt.yticks(rotation=0)
plt.tight_layout()
plt.show()
#plt.savefig("fig_all_tick.png")

