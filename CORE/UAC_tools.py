import os, sys

# ------------------------------------------------------------------------------------------------------------------
def mkdir(folder):

    if not os.path.isdir(folder):
        os.mkdir(folder)

fh = open('log.txt', 'w')
fh.write(sys.argv[-2])
mkdir(sys.argv[-2])
