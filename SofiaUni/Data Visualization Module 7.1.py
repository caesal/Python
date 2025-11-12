import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, FancyArrowPatch

#7.1
fig = plt.figure(figsize=(10, 6))
ax = plt.gca()
ax.set_xlim(0, 10)
ax.set_ylim(0, 6)
ax.axis("off")

def box(x, y, w, h, text):
    ax.add_patch(Rectangle((x, y), w, h, fill=False, linewidth=2))
    ax.text(x + w/2, y + h/2, text, ha="center", va="center", fontsize=10, wrap=True)

def arrow(x1, y1, x2, y2):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="->", mutation_scale=12, linewidth=1.5))

# row one
box(0.5, 3.8, 2.6, 1.0, "Discover program\nSofia University site")
box(4.0, 3.8, 2.6, 1.0, "Create or sign in\nstudent account")
box(7.5, 3.8, 2.6, 1.0, "Search courses\nfilter term and modality")
arrow(3.1, 4.3, 4.0, 4.3); arrow(6.6, 4.3, 7.5, 4.3)

# row two
box(0.5, 2.0, 2.6, 1.0, "Check prerequisites\nand seat availability")
box(4.0, 2.0, 2.6, 1.0, "Consult advisor\noptional chat or email")
box(7.5, 2.0, 2.6, 1.0, "Add class to cart\nreview conflicts")
arrow(3.1, 2.5, 4.0, 2.5); arrow(6.6, 2.5, 7.5, 2.5)

# row three
box(0.5, 0.2, 2.6, 1.0, "Register and submit\naccept policies")
box(4.0, 0.2, 2.6, 1.0, "Tuition payment\nor deferral setup")
box(7.5, 0.2, 2.6, 1.0, "Confirmation email\ncalendar invite and LMS")
arrow(3.1, 0.7, 4.0, 0.7); arrow(6.6, 0.7, 7.5, 0.7)

# down links between rows
for x in [1.8, 5.3, 8.8]:
    arrow(x, 3.8, x, 3.0)
    arrow(x, 2.0, x, 1.2)

plt.title("Customer Journey Map: Student registration at Sofia University")
plt.show()
