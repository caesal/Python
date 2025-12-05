import matplotlib.pyplot as plt

# Dataset 2: Biosyn Dinosaur Investment (2022)
genera = [
    "Allosaurus","Anatosaurus","Brachiosaurus","Camptosaurus",
    "Diplodocus","Iguanodon","Protoceratops","Stegosaurus",
    "Triceratops","Tyrannosaurus"
]

investment_22 = [5, 3, 4, 6, 5, 3, 4, 4, 6, 9]  # in million dollars
total_investment = sum(investment_22)
benchmark = 45

plt.figure(figsize=(8, 2))

# Actual investment bar
plt.barh(0, total_investment, color="tab:blue")

# Benchmark line
plt.axvline(benchmark, color="red", linewidth=2)

# Clean formatting
plt.yticks([])
plt.xlabel("Total quarterly investment in million dollars")
plt.title("Biosyn Genetics Quarterly Investment Bullet Graph")

# Labels using matching colors
plt.text(total_investment + 1, 0, f"Actual {total_investment} M",
         va="center", color="tab:blue")

plt.text(benchmark + 1, 0.2, f"Benchmark {benchmark} M",
         color="red")

# plt.tight_layout()
# plt.show()

import matplotlib.pyplot as plt

# Dataset 2  Biosyn Dinosaur Investment 2022
genera = [
    "Allosaurus","Anatosaurus","Brachiosaurus","Camptosaurus",
    "Diplodocus","Iguanodon","Protoceratops","Stegosaurus",
    "Triceratops","Tyrannosaurus"
]

investment_22 = [5, 3, 4, 6, 5, 3, 4, 4, 6, 9]   # in million dollars
benchmark = 5                                     # benchmark level

plt.figure(figsize=(10, 5))

# Actual line in blue
plt.plot(genera, investment_22, marker="o", linewidth=2,
         color="tab:blue", label="Investment 2022")

# Benchmark line in black
plt.axhline(benchmark, color="red", linewidth=2, label="Benchmark 5 M")

# Text label for benchmark in black
plt.text(len(genera) - 0.8, benchmark + 0.3, "Benchmark 5 M",
         color="black")

plt.title("Investment by Genus with Benchmark Comparison")
plt.xlabel("Genus")
plt.ylabel("Investment in million dollars")
plt.xticks(rotation=30, ha="right")
plt.legend()

# plt.tight_layout()
# plt.show()

import plotly.graph_objects as go

# Dataset 4  Dinosaur Development Status
status = [
    "Not Started",
    "Genetic Extraction",
    "Fertilisation",
    "Embryo Development",
    "Genetic Extraction",
    "Genetic Extraction",
    "Fertilisation",
    "Genetic Extraction",
    "Embryo Development",
    "Embryo Development",
]

genera = [
    "Allosaurus",
    "Anatosaurus",
    "Brachiosaurus",
    "Camptosaurus",
    "Diplodocus",
    "Iguanodon",
    "Protoceratops",
    "Stegosaurus",
    "Triceratops",
    "Tyrannosaurus",
]

hatchlings = [2, 1, 2, 2, 3, 1, 2, 1, 1, 1]

# compute total hatchlings per status, so labels show numbers
status_totals = {}
for st, val in zip(status, hatchlings):
    status_totals[st] = status_totals.get(st, 0) + val

unique_status = list(dict.fromkeys(status))

status_labels = [f"{st} ({status_totals[st]})" for st in unique_status]
genus_labels = [f"{gen} ({val})" for gen, val in zip(genera, hatchlings)]

labels = status_labels + genus_labels

source = []
target = []
values = []

for st, gen, val in zip(status, genera, hatchlings):
    s_idx = labels.index(next(l for l in status_labels if l.startswith(st)))
    t_idx = labels.index(next(l for l in genus_labels if l.startswith(gen)))
    source.append(s_idx)
    target.append(t_idx)
    values.append(val)

fig = go.Figure(data=[go.Sankey(
    node=dict(
        pad=15,
        thickness=20,
        line=dict(width=0.5, color="black"),
        label=labels  # each node shows its count
        # no explicit color, plotly will assign a palette
    ),
    link=dict(
        source=source,
        target=target,
        value=values
        # no explicit link color, plotly will keep them readable
    )
)])

fig.update_layout(
    title_text="Sankey Diagram   Development Status to Genus",
    font_size=10
)

# fig.show()

import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, FancyArrowPatch

fig, ax = plt.subplots(figsize=(10, 5))
ax.set_xlim(0, 10)
ax.set_ylim(0, 5)
ax.axis("off")

def add_box(x, y, w, h, text, facecolor):
    rect = Rectangle((x, y), w, h, linewidth=1.5,
                     edgecolor="black", facecolor=facecolor)
    ax.add_patch(rect)
    ax.text(x + w / 2, y + h / 2, text,
            ha="center", va="center", fontsize=9, wrap=True)

def add_arrow(x1, y1, x2, y2):
    arr = FancyArrowPatch((x1, y1), (x2, y2),
                          arrowstyle="->", mutation_scale=12,
                          linewidth=1.5, color="black")
    ax.add_patch(arr)

# Boxes for investor journey
add_box(0.5, 3.2, 2.5, 0.9,
        "Investor discovers\nBiosyn in media or reports",
        facecolor="lightgray")

add_box(3.5, 3.2, 2.5, 0.9,
        "Visits Biosyn site\nreviews project portfolio",
        facecolor="lightblue")

add_box(6.5, 3.2, 2.5, 0.9,
        "Downloads quarterly report\nand key visuals",
        facecolor="lightgreen")

add_box(0.5, 1.6, 2.5, 0.9,
        "Contacts investor relations\nasks for clarification",
        facecolor="lightyellow")

add_box(3.5, 1.6, 2.5, 0.9,
        "Reviews risk and return\ncompares with other options",
        facecolor="lightpink")

add_box(6.5, 1.6, 2.5, 0.9,
        "Decides on investment\nsize and time frame",
        facecolor="lightcoral")

add_box(3.5, 0.2, 2.5, 0.9,
        "Confirms commitment\nand receives follow up updates",
        facecolor="lightcyan")

# Arrows top row
add_arrow(3.0, 3.65, 3.5, 3.65)
add_arrow(6.0, 3.65, 6.5, 3.65)

# Arrows down
add_arrow(1.75, 3.2, 1.75, 2.5)
add_arrow(7.75, 3.2, 7.75, 2.5)

# Arrows middle row
add_arrow(3.0, 2.05, 3.5, 2.05)
add_arrow(6.0, 2.05, 6.5, 2.05)

# Arrow to final step
add_arrow(4.75, 1.6, 4.75, 1.1)

# plt.title("Customer Journey Map  Investor interaction with Biosyn Genetics")
# plt.tight_layout()
# plt.show()


import matplotlib.pyplot as plt
import numpy as np

# Dataset 1  Brain mass and body mass
groups = [
    "Primates", "Cetacea", "Primates", "Marsupials",
    "Aves", "Proboscidea", "Cetacea", "Marsupials",
    "Fish", "Aves", "Fish", "Reptiles",
    "Carnosaur", "Ornithopod", "Sauropod",
    "Ornithopod", "Sauropod"
]

species = [
    "Homo sapiens",
    "Porpoise",
    "Chimpanzee",
    "Baboon",
    "Crow",
    "Elephant",
    "Blue Whale",
    "Opossum",
    "Goldfish",
    "Ostrich",
    "Tuna",
    "Alligator",
    "Allosaurus",
    "Anatosaurus",
    "Brachiosaurus",
    "Camptosaurus",
    "Diplodocus"
]

brain_mass_g = [
    1420, 1735, 440, 175,
    9.3, 5712, 6800, 4.8,
    0.10, 42.11, 3.09, 14.08,
    167.5, 150, 154.5, 23, 50
]

body_mass_g = [
    83000, 142430, 56690, 19510,
    337, 6654000, 59200000, 1147,
    9.5, 123000, 5210, 205000,
    2300000, 3400000, 87000000, 400000, 11700000
]

# Mark modern animals versus dinosaurs
is_dino = [
    g in ["Carnosaur", "Ornithopod", "Sauropod"]
    for g in groups
]

brain = np.array(brain_mass_g)
body = np.array(body_mass_g)

plt.figure(figsize=(8, 6))

# Modern animals in blue
plt.scatter(
    body[~np.array(is_dino)],
    brain[~np.array(is_dino)],
    label="Modern animals",
    color="tab:blue"
)

# Dinosaurs in orange
plt.scatter(
    body[np.array(is_dino)],
    brain[np.array(is_dino)],
    label="Dinosaurs",
    color="tab:orange"
)

plt.xscale("log")
plt.yscale("log")

plt.xlabel("Body mass in grams  log scale")
plt.ylabel("Brain mass in grams  log scale")
plt.title("Brain mass versus body mass  Modern animals and dinosaurs")
plt.legend()

# plt.tight_layout()
# plt.show()

import matplotlib.pyplot as plt
import numpy as np

# Dataset 2  Investment 2022 and projected 2023
genera = [
    "Allosaurus","Anatosaurus","Brachiosaurus","Camptosaurus",
    "Diplodocus","Iguanodon","Protoceratops","Stegosaurus",
    "Triceratops","Tyrannosaurus"
]

inv_22 = [5, 3, 4, 6, 5, 3, 4, 4, 6, 9]
inv_23 = [1, 2, 3, 4, 2, 1, 2, 1, 4, 5]

x = np.arange(len(genera))
width = 0.5  # Overlap width

plt.figure(figsize=(10, 5))

# 2022 bar in blue
plt.bar(x, inv_22, width, label="2022", color="tab:blue", alpha=0.7)

# 2023 bar in orange, overlapping
plt.bar(x, inv_23, width, label="2023 projected", color="tab:orange", alpha=0.7)

plt.xticks(x, genera, rotation=30, ha="right")
plt.ylabel("Investment in million dollars")
plt.title("Overlapping Bar Chart  Investment 2022 vs 2023 projection")

# plt.legend()
# plt.tight_layout()
# plt.show()

import matplotlib.pyplot as plt
import numpy as np

# Dataset 3  Survey rankings (1 = highest interest)
genera = [
    "Allosaurus","Anatosaurus","Brachiosaurus","Camptosaurus",
    "Diplodocus","Iguanodon","Protoceratops","Stegosaurus",
    "Triceratops","Tyrannosaurus"
]

rankings = [3, 4, 2, 5, 5, 3, 4, 4, 2, 1]

# Convert to diverging scale around midpoint
midpoint = 3  # center of interest
diverging = np.array(rankings) - midpoint

colors = ["tab:blue" if val < 0 else "tab:red" for val in diverging]

plt.figure(figsize=(10, 6))

y = np.arange(len(genera))

plt.barh(y, diverging, color=colors)

# Center line
plt.axvline(0, color="black", linewidth=1)

plt.yticks(y, genera)
plt.xlabel("Interest relative to midpoint")
plt.title("Diverging Bar Chart  Public interest in dinosaur genera")

# Add labels at end of bars
for i, val in enumerate(diverging):
    plt.text(val + 0.1 if val > 0 else val - 0.3,
             i,
             f"{rankings[i]}",
             va="center",
             color="black")

# plt.tight_layout()
# plt.show()

import matplotlib.pyplot as plt
import numpy as np

# Dataset 2  Investment 2023 projection
genera = [
    "Allosaurus","Anatosaurus","Brachiosaurus","Camptosaurus",
    "Diplodocus","Iguanodon","Protoceratops","Stegosaurus",
    "Triceratops","Tyrannosaurus"
]

inv_23 = [1, 2, 3, 4, 2, 1, 2, 1, 4, 5]   # in million dollars

# Sort for clearer dot plot
sorted_indices = np.argsort(inv_23)
genera_sorted = [genera[i] for i in sorted_indices]
inv_23_sorted = [inv_23[i] for i in sorted_indices]

y = np.arange(len(genera_sorted))

plt.figure(figsize=(8, 6))

plt.scatter(inv_23_sorted, y, color="tab:blue")

plt.yticks(y, genera_sorted)
plt.xlabel("Projected investment in million dollars  2023")
plt.title("Dot Plot  Projected Dinosaur Investment 2023 by Genus")

# Add labels next to dots
for x_val, y_val in zip(inv_23_sorted, y):
    plt.text(x_val + 0.1, y_val, f"{x_val}", va="center")

plt.tight_layout()
plt.show()
