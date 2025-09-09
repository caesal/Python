# Machine Learning Project Assignment Visualization.py

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from sklearn.datasets import load_breast_cancer
from sklearn.feature_selection import f_classif
from sklearn.preprocessing import StandardScaler

# 1) Load dataset
data = load_breast_cancer()
X = pd.DataFrame(data.data, columns=data.feature_names)
y = pd.Series(data.target, name="class")  # 0 = malignant, 1 = benign

# 2) Pick top 4 features by ANOVA F-test
F, p = f_classif(X, y)
top_idx = np.argsort(F)[-4:][::-1]  # best 4 features
top_features = X.columns[top_idx].tolist()
print("Top 4 selected features:", top_features)

# 3) Standardize selected features (for visualization only)
scaler = StandardScaler()
X_top = pd.DataFrame(
    scaler.fit_transform(X[top_features]),
    columns=top_features
)

# 4) Split by class
X_mal = X_top[y == 0]  # malignant
X_ben = X_top[y == 1]  # benign

# 5) Build pair-plot style scatter matrix
n = len(top_features)
fig, axes = plt.subplots(n, n, figsize=(10, 10))
fig.suptitle("Breast Cancer Diagnostic — Pairwise Feature Matrix\nTop 4 features by ANOVA F-test (standardized)", y=0.93)

for i, fi in enumerate(top_features):
    for j, fj in enumerate(top_features):
        ax = axes[i, j]
        if i == j:
            # Diagonal: histograms
            ax.hist(X_ben[fj], bins=30, alpha=0.6, label="benign", color="orange")
            ax.hist(X_mal[fj], bins=30, alpha=0.6, label="malignant", color="blue")
        else:
            # Off-diagonal: scatter
            ax.scatter(X_ben[fj], X_ben[fi], s=10, label="benign", alpha=0.6, color="orange", marker="o")
            ax.scatter(X_mal[fj], X_mal[fi], s=10, label="malignant", alpha=0.6, color="blue", marker="x")
        if i == n - 1:
            ax.set_xlabel(fj, rotation=45, ha="right")
        else:
            ax.set_xticks([])
        if j == 0:
            ax.set_ylabel(fi)
        else:
            ax.set_yticks([])

# Single legend for the whole figure
handles, labels = axes[0, 0].get_legend_handles_labels()
fig.legend(handles[:2], labels[:2], loc="lower center", ncol=2, frameon=False)

plt.tight_layout(rect=[0, 0.04, 1, 0.92])

# 6) Save outputs
png_path = "breast_cancer_pairplot.png"
pdf_path = "breast_cancer_pairplot.pdf"
plt.savefig(png_path, dpi=180, bbox_inches="tight")
plt.savefig(pdf_path, dpi=180, bbox_inches="tight")

plt.show()

print(f"Figures saved as:\n- {png_path}\n- {pdf_path}")
