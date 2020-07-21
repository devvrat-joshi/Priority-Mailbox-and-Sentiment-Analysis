import antispam, sys
x = " ".join(sys.argv[2:])
print(antispam.score(x))