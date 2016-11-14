import os, sys, zipfile, shutil
import _winreg
import win32com.client
import logging


INST_FOLDER = sys.argv[1].replace('%20', ' ')
RHINO_FOLDER = sys.argv[2].replace('%20', ' ')
PACKAGE = sys.argv[3].replace('%20', ' ')

LIB_FOLDER = r'Plug-ins\IronPython\Lib\\'
CONFIG_FILE = 'MW_printer_global_parameter.py'
TOOLBAR_FOLDER = r'Plug-ins\Toolbars\\'
TOOLBAR_FILE = 'MW3DPrint_TB.rui'


# ----------------------------------------------------------------------------------------------------------------------
def mkdir(INST_FOLDER):

    try:
        inst_folder_split = INST_FOLDER.split('\\')

        folder_str = ''
        for folder in inst_folder_split:
            folder_str = folder_str + folder + '\\'
            if not os.path.isdir(folder_str):
                os.mkdir(folder_str)
    except:
        MWLOG.exception('CREATE FOLDER')

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
        except Exception as e:
            MWLOG.exception('EXTRACTING')

    fh.close()


# ----------------------------------------------------------------------------------------------------------------------
def copy_conf_file(INST_FOLDER, RHINO_FOLDER):
    try:
        shutil.move(INST_FOLDER + CONFIG_FILE, RHINO_FOLDER + LIB_FOLDER + CONFIG_FILE)
    except Exception as e:
        MWLOG.exception('COPY CONFIG FILE')


# ----------------------------------------------------------------------------------------------------------------------
def copy_tool_bar(INST_FOLDER, RHINO_FOLDER):
    try:
        shutil.move(INST_FOLDER + TOOLBAR_FILE, RHINO_FOLDER + TOOLBAR_FOLDER + TOOLBAR_FILE)
    except Exception as e:
        MWLOG.exception('COPY TOOL BAR')

# ----------------------------------------------------------------------------------------------------------------------
def adapt_abs_folders(INST_FOLDER, RHINO_FOLDER):
    _ABS_FOLDER_PLUGIN = INST_FOLDER + 'PY'
    _ABS_FOLDER_SCRIPTS = INST_FOLDER + r'PY\RhinoInterface'
    _ABS_FOLDER_CORE = INST_FOLDER

    try:

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

    except Exception as e:
        MWLOG.exception('ADAPT ABS FODLER')

# ----------------------------------------------------------------------------------------------------------------------
def create_reg_key():
    try:
        reg = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, 'Software')
        reg_mw = _winreg.CreateKey(reg, 'MW3DPrinting')
        _winreg.SetValue(reg_mw, 'CorePath', 1, INST_FOLDER)
        _winreg.SetValue(reg_mw, 'ConfigFile', 1, RHINO_FOLDER + LIB_FOLDER + CONFIG_FILE)
        # _winreg.CloseKey(_winreg.HKEY_LOCAL_MACHINE)
    except Exception as e:
        MWLOG.exception('REG KEY')


# ----------------------------------------------------------------------------------------------------------------------
def create_startmenu_shortcut():

    try:
        shortcut_group_name = "Moduleworks"
        shortcut_name = "Visit ModuleWorks website"
        shortcut_target = "http://www.moduleworks.com"

        sh = win32com.client.Dispatch("WScript.Shell")
        p = sh.SpecialFolders("AllUsersPrograms")
        assert(os.path.isdir(p))
        p = os.path.join(p, shortcut_group_name)

        if not os.path.isdir(p):
            os.makedirs(p)

        lnk = sh.CreateShortcut(os.path.join(p, shortcut_name + ".lnk"))
        lnk.TargetPath = shortcut_target
        lnk.Save()

    except Exception as e:
        MWLOG.exception('CREATING SHORTCUT')



MWLOG = logging.getLogger('MWSETUP')
MWLOG.info('Starting installation ...')

MWLOG.info('Creating installation folder')
mkdir(INST_FOLDER)
MWLOG.info('Extracting package')
extract(INST_FOLDER, PACKAGE)
MWLOG.info('Copy config file')
copy_conf_file(INST_FOLDER, RHINO_FOLDER)
MWLOG.info('Copy tool bar')
copy_tool_bar(INST_FOLDER, RHINO_FOLDER)
MWLOG.info('Adapting config file')
adapt_abs_folders(INST_FOLDER, RHINO_FOLDER)
MWLOG.info('Creating registry entry')
create_reg_key()
MWLOG.info('Creating start menu')
create_startmenu_shortcut()
