import os, sys

# ------------------------------------------------------------------------------------------------------------------
def mkdir(folder):

    if not os.path.isdir(folder):
        os.mkdir(folder)

print 'RUNNING EXTERN'
mkdir(r'C:\Program Files\MW3DPrinting\\')
