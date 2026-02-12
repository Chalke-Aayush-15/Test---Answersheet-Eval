from datetime import datetime, timedelta

def time_to_seconds(time_str):
    h, m, s = map(int, time_str.split(":"))
    return h * 3600 + m * 60 + s

def seconds_to_time(seconds):
    seconds = seconds % (24 * 3600)   # handle overflow
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

print("=== Berkeley Clock Synchronization Algorithm ===")

# Input number of machines
n = int(input("Enter number of machines: "))

times = []
for i in range(n):
    t = input(f"Enter time for Machine {i} (HH:MM:SS): ")
    times.append(time_to_seconds(t))

# Input master ID
master_id = int(input("Enter Master Machine ID (0 to N-1): "))

print("\nStep 1: Master polls all slave machines for their time.")

master_time = times[master_id]
offsets = []

print("\nStep 2: Calculating time differences (offsets) from Master:\n")

for i in range(n):
    offset = times[i] - master_time
    offsets.append(offset)
    print(f"Machine {i} Offset = {offset} seconds")

print("\nStep 3: Calculating Average Offset.")

avg_offset = sum(offsets) / n
print("Average Offset =", avg_offset, "seconds")

print("\nStep 4: Sending offset adjustment to each machine.")

new_times = []

for i in range(n):
    adjustment = avg_offset - offsets[i]
    print(f"Machine {i} Adjustment = {adjustment} seconds")
    new_time = times[i] + adjustment
    new_times.append(new_time)

print("\nStep 5: New Synchronized Time at each machine:\n")

for i in range(n):
    print(f"Machine {i} New Time = {seconds_to_time(int(new_times[i]))}")

print("\nConclusion: Berkeley Clock Synchronization Successfully Implemented.")
