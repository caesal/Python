import matplotlib.pyplot as plt
import numpy as np

#7.2
fig = plt.figure(figsize=(8, 5))

steps = ["Discover", "Sign in", "Search", "Prereq", "Advisor", "Add cart", "Register", "Payment", "Confirm"]
channels = ["Desktop", "Mobile", "Helpdesk", "Email", "LMS"]

# sample intensity scores
data = np.array([
    [2, 3, 4, 3, 1, 2, 2, 3, 1],
    [3, 4, 5, 2, 1, 3, 2, 4, 2],
    [1, 2, 2, 2, 4, 3, 2, 1, 1],
    [1, 2, 3, 2, 3, 2, 2, 2, 4],
    [2, 2, 2, 3, 3, 2, 3, 2, 3],
])

im = plt.imshow(data, aspect="auto")
plt.colorbar(im, label="Relative intensity")
plt.xticks(range(len(steps)), steps, rotation=30, ha="right")
plt.yticks(range(len(channels)), channels)
plt.title("Heat Map: Relative intensity by channel and step")
plt.xlabel("Journey steps")
plt.ylabel("Channels")
plt.tight_layout()
plt.show()

#7.3
fig = plt.figure(figsize=(9, 5))

years = np.arange(2018, 2026)
undergrad = np.array([30, 28, 26, 25, 23, 22, 20, 19])
graduate = np.array([25, 26, 27, 28, 28, 29, 30, 31])
cont_ed  = 100 - undergrad - graduate

plt.stackplot(years, undergrad, graduate, cont_ed, labels=["Undergraduate", "Graduate", "Continuing ed"])
plt.legend(loc="upper left")
plt.title("Histomap style share over time")
plt.xlabel("Year")
plt.ylabel("Percent of total enrollment share")
plt.tight_layout()
plt.show()
