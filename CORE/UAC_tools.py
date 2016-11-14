import os, sys, zipfile, shutil
import _winreg

INST_FOLDER = sys.argv[1].replace('%20', ' ')
RHINO_FOLDER = sys.argv[2].replace('%20', ' ')
PACKAGE = sys.argv[3].replace('%20', ' ')

LIB_FOLDER = r'Plug-ins\IronPython\Lib\\'
CONFIG_FILE = 'MW_printer_global_parameter.py'
TOOLBAR_FOLDER = r'Plug-ins\Toolbars\\'
TOOLBAR_FILE = 'MW3DPrint_TB.rui'


# ----------------------------------------------------------------------------------------------------------------------
def mkdir(INST_FOLDER):

    inst_folder_split = INST_FOLDER.split('\\')

    folder_str = ''
    for folder in inst_folder_split:
        folder_str = folder_str + folder + '\\'
        if not os.path.isdir(folder_str):
            os.mkdir(folder_str)


# ----------------------------------------------------------------------------------------------------------------------
def extract(INST_FOLDER, PACKAGE):
    fh = open(PACKAGE, 'rb')
    z = zipfile.ZipFile(fh)
    for name in z.namelist():
        try:
            print name
            if str(name)[-1] == '/':

                folders = name.split('/')
                folder_str = ''
                for folder in folders:
                    folder_str = folder_str + folder + '\\'
                    print 'Creating: ' + INST_FOLDER + folder_str
                    try:
                        os.mkdir(INST_FOLDER + folder_str)
                    except:
                        pass

            else:
                outfile = open(INST_FOLDER + name, 'wb')
                outfile.write(z.read(name))
                outfile.close()
        except:
            print name + ' failed'

    fh.close()


# ----------------------------------------------------------------------------------------------------------------------
def copy_conf_file(INST_FOLDER, RHINO_FOLDER):
    shutil.move(INST_FOLDER + CONFIG_FILE, RHINO_FOLDER + LIB_FOLDER + CONFIG_FILE)


# ----------------------------------------------------------------------------------------------------------------------
def copy_tool_bar(INST_FOLDER, RHINO_FOLDER):
    shutil.move(INST_FOLDER + TOOLBAR_FILE, RHINO_FOLDER + TOOLBAR_FOLDER + TOOLBAR_FILE)


# ----------------------------------------------------------------------------------------------------------------------
def adapt_abs_folders(INST_FOLDER, RHINO_FOLDER):
    _ABS_FOLDER_PLUGIN = INST_FOLDER + 'PY'
    _ABS_FOLDER_SCRIPTS = INST_FOLDER + r'PY\RhinoInterface'
    _ABS_FOLDER_CORE = INST_FOLDER

    fh = open(RHINO_FOLDER + LIB_FOLDER + CONFIG_FILE, 'r')
    content = fh.readlines()
    for line, j in zip(content, range(len(content))):
        if line.find('_ABS_FOLDER_PLUGIN') != -1:
            line_split = line.split('=')
            line_split_split = line_split[1].split("'")
            line_split_split[1] = _ABS_FOLDER_PLUGIN
            line_split[1] = "'".join(line_split_split)
            line = "=".join(line_split)
            content[j] = line

        if line.find('_ABS_FOLDER_SCRIPTS') != -1:
            line_split = line.split('=')
            line_split_split = line_split[1].split("'")
            line_split_split[1] = _ABS_FOLDER_SCRIPTS
            line_split[1] = "'".join(line_split_split)
            line = "=".join(line_split)
            content[j] = line

        if line.find('_ABS_FOLDER_CORE') != -1:
            line_split = line.split('=')
            line_split_split = line_split[1].split("'")
            line_split_split[1] = _ABS_FOLDER_CORE[:-1]
            line_split[1] = "'".join(line_split_split)
            line = "=".join(line_split)
            content[j] = line

    fh.close()

    fh = open(RHINO_FOLDER + LIB_FOLDER + CONFIG_FILE, 'w')
    for line in content:
        fh.write(line)

    fh.close()


# ----------------------------------------------------------------------------------------------------------------------
def create_reg_key():
    try:
        reg = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, 'Software')
        reg_mw = _winreg.CreateKey(reg, 'MW3DPrinting')
        _winreg.SetValue(reg_mw, 'CorePath', 1, INST_FOLDER)
        _winreg.SetValue(reg_mw, 'ConfigFile', 1, RHINO_FOLDER + LIB_FOLDER + CONFIG_FILE)
        _winreg.CloseKey(_winreg.HKEY_LOCAL_MACHINE)
    except:
        fh = open('log.txt', 'wa')
        fh.write('[WARNING] Could not create reg keys.')
        fh.close()


mkdir(INST_FOLDER)
extract(INST_FOLDER, PACKAGE)
copy_conf_file(INST_FOLDER, RHINO_FOLDER)
copy_tool_bar(INST_FOLDER, RHINO_FOLDER)
adapt_abs_folders(INST_FOLDER, RHINO_FOLDER)
create_reg_key()
