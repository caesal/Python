"""
Author: Junjie Cheng
"""

import os
import time
from multiprocessing import Process

# -----------------------------
# Task 1: Print Process IDs
# -----------------------------
def child_task_print_pid():
    print(f"[Task 1] Child PID: {os.getpid()}  Parent PID seen by child: {os.getppid()}")

def task1():
    print("\n=== Task 1: Print Process IDs ===")
    print(f"[Task 1] Parent PID: {os.getpid()}")
    p = Process(target=child_task_print_pid)
    p.start()
    p.join()
    print("[Task 1] Child finished\n")


# -----------------------------
# Task 2: Run Two Processes in Parallel
# -----------------------------
def print_numbers(name):
    for i in range(1, 6):
        print(f"[Task 2] Process {name} PID {os.getpid()}: {i}")
        time.sleep(1)

def task2():
    print("\n=== Task 2: Two Processes in Parallel ===")
    p1 = Process(target=print_numbers, args=("A",))
    p2 = Process(target=print_numbers, args=("B",))

    start_time = time.time()
    p1.start()
    p2.start()

    p1.join()
    p2.join()
    end_time = time.time()

    elapsed = end_time - start_time
    print(f"[Task 2] Both processes are done in about {elapsed:.2f} seconds\n")


# -----------------------------
# Task 3: Process Synchronization with join()
# -----------------------------
def say(msg, delay=0.5):
    time.sleep(delay)
    print(f"[Task 3] {msg} (PID {os.getpid()})")

def task3():
    print("\n=== Task 3: Process Synchronization with join() ===")
    p1 = Process(target=say, args=("Hello", 0.8))
    p2 = Process(target=say, args=("World", 0.3))

    p1.start()
    p2.start()

    # Parent waits for both children to finish before printing the final line
    p1.join()
    p2.join()

    print("[Task 3] All processes finished\n")


if __name__ == "__main__":
    task1()
    task2()
    task3()


# ------------------------------------------------
# Short Explanation (Deliverables)
# ------------------------------------------------
# What a process is:
# A process is a running instance of a program with its own memory space
# and a unique Process ID (PID).
#
# What I observed when running the code:
# - In Task 1, the parent and child have different PIDs.
# - In Task 2, outputs from both processes are interleaved, showing parallel execution.
# - In Task 3, the final line prints only after both children finish.
#
# How process synchronization works with .join():
# The .join() method makes the parent wait until the child finishes.
# By joining all children, the parent ensures it continues only after
# all processes are done.
