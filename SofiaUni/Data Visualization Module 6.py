import matplotlib.pyplot as plt
import pandas as pd
import networkx as nx

# Dataset11
data = {
    "Category": [
        "species", "subspecies", "subspecies", "subspecies", "subspecies",
        "species", "species", "subspecies", "subspecies", "subspecies", "subspecies", "subspecies"
    ],
    "English name": [
        "Common Ostrich", "", "", "", "", "Somali Ostrich", "Greater Rhea", "", "", "", "", ""
    ],
    "Scientific Name": [
        "Struthio camelus", "Struthio camelus camelus", "Struthio camelus syriacus",
        "Struthio camelus massaicus", "Struthio camelus australis", "Struthio molybdophanes",
        "Rhea americana", "Rhea americana araneipes", "Rhea americana americana",
        "Rhea americana nobilis", "Rhea americana albescens", "Rhea americana intermedia"
    ],
    "Order": [
        "Struthioniformes", "Struthioniformes", "Struthioniformes", "Struthioniformes", "Struthioniformes",
        "Struthioniformes", "Rheiformes", "Rheiformes", "Rheiformes", "Rheiformes", "Rheiformes", "Rheiformes"
    ],
    "Family": [
        "Struthionidae (Ostriches)", "Struthionidae (Ostriches)", "Struthionidae (Ostriches)", "Struthionidae (Ostriches)", "Struthionidae (Ostriches)",
        "Struthionidae (Ostriches)", "Rheidae (Rheas)", "Rheidae (Rheas)", "Rheidae (Rheas)", "Rheidae (Rheas)", "Rheidae (Rheas)", "Rheidae (Rheas)"
    ]
}

df = pd.DataFrame(data)

# Diagram 1: Linear timeline style of connections
plt.figure(figsize=(10, 4))
species_positions = {
    "Struthio camelus": 1,
    "Struthio molybdophanes": 2,
    "Rhea americana": 3
}

# Plot species and subspecies along a timeline
for _, row in df.iterrows():
    x = 1 if "camelus" in row["Scientific Name"] else 2 if "molybdophanes" in row["Scientific Name"] else 3
    plt.scatter(x, species_positions.get(row["Scientific Name"].split()[0], species_positions.get("Rhea americana")), color="teal", s=200)
    plt.text(x + 0.05, species_positions.get(row["Scientific Name"].split()[0], 1) + 0.05, row["Scientific Name"], rotation=30, fontsize=8)

plt.yticks([1, 2, 3], ["Common Ostrich", "Somali Ostrich", "Greater Rhea"])
plt.title("Diagram 1: Linear Connection of Species and Subspecies")
plt.xlabel("Timeline / Evolutionary Branch")
plt.ylabel("Species")
plt.grid(True, linestyle="--", alpha=0.6)
plt.tight_layout()
plt.show()

# Diagram 2: Network focusing on scientific naming hierarchy within species
G = nx.DiGraph()
species = "Struthio camelus"
G.add_node(species, label="Common Ostrich")

for sub in df[df["Scientific Name"].str.startswith(species + " ")]["Scientific Name"]:
    G.add_edge(species, sub)

plt.figure(figsize=(8, 6))
pos = nx.spring_layout(G, k=0.6, seed=42)
nx.draw(G, pos, with_labels=True, node_color="lightgreen", node_size=2000, font_size=8, arrows=False)
plt.title("Diagram 2: Scientific Naming Hierarchy within Common Ostrich (Struthio camelus)")
plt.show()
