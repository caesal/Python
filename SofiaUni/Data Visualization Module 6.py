# Two clean visualizations (no saving, only plotting)

import matplotlib.pyplot as plt

# --- Shared data ---
species = [
    ("Struthio camelus", "Common Ostrich", "Struthioniformes", "Struthionidae (Ostriches)"),
    ("Struthio molybdophanes", "Somali Ostrich", "Struthioniformes", "Struthionidae (Ostriches)"),
    ("Rhea americana", "Greater Rhea", "Rheiformes", "Rheidae (Rheas)"),
]

subspecies_map = {
    "Struthio camelus": [
        "Struthio camelus camelus",
        "Struthio camelus syriacus",
        "Struthio camelus massaicus",
        "Struthio camelus australis",
    ],
    "Struthio molybdophanes": [],
    "Rhea americana": [
        "Rhea americana araneipes",
        "Rhea americana americana",
        "Rhea americana nobilis",
        "Rhea americana albescens",
        "Rhea americana intermedia",
    ],
}

# --- Diagram 1: Linear / timeline-style connection ---
fig = plt.figure(figsize=(11, 5))

x_levels = ["Order", "Family", "Species", "Subspecies"]
xpos = {lvl: i for i, lvl in enumerate(x_levels)}
lanes = {sp[0]: idx for idx, sp in enumerate(reversed(species))}

for sp, eng, order, family in species:
    y = lanes[sp]
    plt.plot([xpos["Order"], xpos["Family"], xpos["Species"]], [y, y, y], marker="o", linewidth=2)
    plt.text(xpos["Order"], y + 0.12, order, ha="center", va="bottom", fontsize=10)
    plt.text(xpos["Family"], y + 0.12, family, ha="center", va="bottom", fontsize=10)
    plt.text(xpos["Species"], y + 0.12, f"{sp}\n({eng})", ha="center", va="bottom", fontsize=10)

    subs = sorted(subspecies_map.get(sp, []))
    if subs:
        offsets = [i - (len(subs) - 1) / 2 for i in range(len(subs))]
        for off, sub in zip(offsets, subs):
            y_sub = y + 0.3 * off
            plt.plot([xpos["Species"], xpos["Subspecies"]], [y, y_sub], linestyle="--", linewidth=1)
            plt.plot([xpos["Subspecies"]], [y_sub], marker="o")
            plt.text(xpos["Subspecies"] + 0.02, y_sub, sub, ha="left", va="center", fontsize=9)

plt.xticks(list(xpos.values()), x_levels)
plt.yticks(list(lanes.values()), [f"{sp[0]} lane" for sp in reversed(species)])
plt.title("Linear connection across taxonomy: Order → Family → Species → Subspecies", fontsize=13)
plt.xlabel("Taxonomic level (left → right)")
plt.ylabel("Species lanes")
plt.grid(True, linestyle="--", alpha=0.4)
plt.tight_layout()
plt.show()

# --- Diagram 2: Scientific-name-focused tree (within one species) ---
focus = "Struthio camelus"
children = sorted(subspecies_map[focus])

fig2 = plt.figure(figsize=(7, 5))
root_x, root_y = 0.0, 0.5

if children:
    ys = [0.8 - i * (0.6 / (len(children) - 1 if len(children) > 1 else 1)) for i in range(len(children))]
    for y_sub, sub in zip(ys, children):
        plt.plot([root_x, 0.6], [root_y, y_sub])  # edge
        plt.plot(0.6, y_sub, marker="o")
        short = sub.replace(focus + " ", "")
        plt.text(0.62, y_sub, short, va="center", fontsize=10)

plt.plot(root_x, root_y, marker="o")
plt.text(root_x - 0.02, root_y, focus, ha="right", va="center", fontsize=11)
plt.title("Scientific naming within one species: Struthio camelus", fontsize=13)
plt.axis("off")
plt.tight_layout()
plt.show()
