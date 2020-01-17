import sys
writer = sys.stdout.write
for i in range(0, 1000):
    writer("\n Doing " + str(i))
