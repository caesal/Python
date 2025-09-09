# KNN ROC Curve.py

import matplotlib.pyplot as plt
from sklearn.datasets import load_breast_cancer
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import Pipeline
from sklearn.neighbors import KNeighborsClassifier
from sklearn.metrics import RocCurveDisplay

# 1. Load dataset
data = load_breast_cancer()
X, y = data.data, data.target  # target: 0 = malignant, 1 = benign

# 2. Train-test split
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42, stratify=y
)

# 3. Define BEST KNN model (k=3, Euclidean distance, distance weighting)
best_knn = Pipeline([
    ("scaler", StandardScaler()),
    ("knn", KNeighborsClassifier(n_neighbors=3, metric="euclidean", weights="distance"))
])

# 4. Fit model
best_knn.fit(X_train, y_train)

# 5. Predict probabilities for malignant (label=0)
y_proba = best_knn.predict_proba(X_test)[:, 0]

# 6. Plot ROC curve
RocCurveDisplay.from_predictions(y_test == 0, y_proba, name="KNN (k=3, Euclidean, distance)")
plt.title("ROC Curve for Best KNN Model")
plt.savefig("knn_best_roc_curve.png", dpi=180, bbox_inches="tight")
plt.show()

print("ROC Curve saved as knn_best_roc_curve.png")
