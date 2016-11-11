import os, sys, zipfile, shutil


# ------------------------------------------------------------------------------------------------------------------
def mkdir(inst_folder):
    if not os.path.isdir(inst_folder):
        os.mkdir(inst_folder)


# ------------------------------------------------------------------------------------------------------------------
def extract(inst_folder, package):
    fh = open(package, 'rb')
    z = zipfile.ZipFile(fh)
    for name in z.namelist():
        try:
            print name
            if str(name)[-1] == '/':

                folders = name.split('/')
                folder_str = ''
                for folder in folders:
                    folder_str = folder_str + folder + '\\'
                    print 'Creating: ' + inst_folder + folder_str
                    try:
                        os.mkdir(inst_folder + folder_str)
                    except:
                        pass

            else:
                outfile = open(inst_folder + name, 'wb')
                outfile.write(z.read(name))
                outfile.close()
        except:
            print name + ' failed'

    fh.close()


# ------------------------------------------------------------------------------------------------------------------
def copy_conf_file(inst_folder, rhino_folder):
    config_file = 'MW_printer_global_parameter.py'
    lib_folder = r'Plug-ins\IronPython\Lib\\'
    shutil.move(inst_folder + config_file, rhino_folder + lib_folder + config_file)


# ------------------------------------------------------------------------------------------------------------------
def adapt_abs_folders():
    pass


mkdir(sys.argv[-4].replace('%20', ' '))
extract(sys.argv[-4].replace('%20', ' '), sys.argv[-2].replace('%20', ' '))
copy_conf_file(sys.argv[-4].replace('%20', ' '), sys.argv[-3].replace('%20', ' '))
