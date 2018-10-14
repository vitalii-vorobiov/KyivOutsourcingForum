with open("user.txt") as f:

    for line in f.readlines():
        if "vorobyov@ucu.edu.ua" in line:
            print(True)
        else:
            print(False)