# testing.py6
import math

x1 = 10
y1 = 10

x2 = 20
y2 = 20

delta_x = x1 - x2
delta_y = y1 - y2

distance_0 = math.sqrt(delta_x**2 + delta_y**2) * 12
distance_1 = math.hypot(delta_x, delta_y) * 12

print(distance_0)
print(distance_1)


dx = delta_x
dy = delta_y

cen_w = (dx * dx + dy * dy) ** 0.5 * 12.0
cen_w1 = math.hypot(dx, dy) * 12

print(cen_w)
print(cen_w1)

print("# Big text")
