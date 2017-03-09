import os, sys, zipfile, shutil, time
import _winreg
import win32com.client
import logging

print 'Execution string ' + str(sys.argv)

if len(sys.argv) > 5:
    INST_FOLDER = sys.argv[1].replace('%20', ' ')
    RHINO_FOLDER = sys.argv[2].replace('%20', ' ')
    PACKAGE = sys.argv[3].replace('%20', ' ')
    SETUP_FOLDER = sys.argv[4].replace('%20', ' ')
    SETUP_ROUTINE = sys.argv[5].replace('%20', ' ') # install or uninstall
else:
    print 'Folder destinations are missing.'
    print '1. -> Installation folder'
    print '2. -> Rhino folder'
    print '3. -> Package folder + package'
    print '4. -> Setup folder'
    print '5. -> install or uninstall'

print sys.argv

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
def remove_folders(folders):
    try:
        for folder in folders:
            shutil.rmtree(folder)
    except Exception as e:
        MWLOG.exception('REMOVE FOLDER')

# ----------------------------------------------------------------------------------------------------------------------
def extract(INST_FOLDER, PACKAGE):
    try:
        fh = open(PACKAGE, 'rb')
        z = zipfile.ZipFile(fh)
        for name in z.namelist():
            print name
            MWLOG.info(' -> ' + name)
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

# ----------------------------------------------------------------------------------------------------------------------
def copy_conf_file(INST_FOLDER, RHINO_FOLDER):
    try:
        shutil.move(INST_FOLDER + CONFIG_FILE, RHINO_FOLDER + LIB_FOLDER + CONFIG_FILE)
    except Exception as e:
        MWLOG.exception('COPY CONFIG FILE')

def remove_conf_file(RHINO_FOLDER):
    try:
        os.remove(RHINO_FOLDER + LIB_FOLDER + CONFIG_FILE)
    except Exception as e:
        MWLOG.exception('REMOVING CONFIG FILE')

# ----------------------------------------------------------------------------------------------------------------------
def copy_tool_bar(INST_FOLDER, RHINO_FOLDER):
    try:
        shutil.move(INST_FOLDER + TOOLBAR_FILE, RHINO_FOLDER + TOOLBAR_FOLDER + TOOLBAR_FILE)
    except Exception as e:
        MWLOG.exception('COPY TOOL BAR')

# ----------------------------------------------------------------------------------------------------------------------
def remove_toolbar(tool_bar_file):
    try:
        os.remove(tool_bar_file)
    except Exception as e:
        MWLOG.exception('REMOVING TOOLBAR')

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
        reg = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, r'Software\MW')
        reg_mw = _winreg.CreateKey(reg, '3DPrinting')
        _winreg.SetValue(reg_mw, 'CorePath', 1, INST_FOLDER)
        _winreg.SetValue(reg_mw, 'ConfigFile', 1, RHINO_FOLDER + LIB_FOLDER + CONFIG_FILE)
        # _winreg.CloseKey(_winreg.HKEY_LOCAL_MACHINE)
    except Exception as e:
        MWLOG.exception('REG KEY')

# ----------------------------------------------------------------------------------------------------------------------
def remove_reg_key():
    try:
        root_key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, r'Software\MW')
        _winreg.DeleteKey(root_key, '3DPrinting\CorePath')
        _winreg.DeleteKey(root_key, '3DPrinting\ConfigFile')
        _winreg.DeleteKey(root_key, '3DPrinting')
    except:
        MWLOG.exception('REMOVE REG KEY')

# ----------------------------------------------------------------------------------------------------------------------
def create_startmenu_shortcut():

    try:
        shortcut_group_name = 'Moduleworks'
        shortcut_name = ['Visit ModuleWorks website', 'Rhino MW3D Printing']
        shortcut_target = ['http://www.moduleworks.com',
                           str(RHINO_FOLDER) + r'\System\Rhino.exe']
        shortcut_args = ['""', str(RHINO_FOLDER) + r'Plug-ins\Toolbars\MW3DPrint_TB.rui']

        sh = win32com.client.Dispatch("WScript.Shell")
        p = sh.SpecialFolders("AllUsersPrograms")
        assert(os.path.isdir(p))
        p = os.path.join(p, shortcut_group_name)

        if not os.path.isdir(p):
            os.makedirs(p)

        for i in range(len(shortcut_name)):
            lnk = sh.CreateShortcut(os.path.join(p, shortcut_name[i] + '.lnk'))
            lnk.TargetPath = shortcut_target[i]
            lnk.Arguments = shortcut_args[i]
            lnk.Save()

    except Exception as e:
        MWLOG.exception('CREATING SHORTCUT')

# ----------------------------------------------------------------------------------------------------------------------
def remove_startmenu_shortcut():

    try:
        shortcut_group_name = 'Moduleworks'
        shortcut_name = ['Visit ModuleWorks website', 'Rhino MW3D Printing']
        shortcut_target = ['http://www.moduleworks.com',
                           str(RHINO_FOLDER) + r'\System\Rhino.exe']
        shortcut_args = ['""', str(RHINO_FOLDER) + r'Plug-ins\Toolbars\MW3DPrint_TB.rui']

        sh = win32com.client.Dispatch("WScript.Shell")
        p = sh.SpecialFolders("AllUsersPrograms")
        assert (os.path.isdir(p))
        p = os.path.join(p, shortcut_group_name)

        if not os.path.isdir(p):
            os.makedirs(p)

        for i in range(len(shortcut_name)):
            lnk = sh.CreateShortcut(os.path.join(p, shortcut_name[i] + '.lnk'))
            lnk.TargetPath = shortcut_target[i]
            lnk.Arguments = shortcut_args[i]
            lnk.Save()

    except Exception as e:
        MWLOG.exception('CREATING SHORTCUT')

# ----------------------------------------------------------------------------------------------------------------------
def install():
    MWLOG.info('---------------------------------------------')
    MWLOG.info('MW installer start: ' + time.ctime())
    MWLOG.info('---------------------------------------------')

    if len(sys.argv) > 4:
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
        MWLOG.info('DONE')
    else:
        print 'Folder destinations are missing.'
        MWLOG.error('Folder destinations are missing. Add sys arguments.')

    # move logfile
    try:
        shutil.move(r'MWLOG_UAC.log', SETUP_FOLDER + r'\MWLOG_UAC.log')
    except:
        pass


# ----------------------------------------------------------------------------------------------------------------------
def uninstall():
    MWLOG.info('---------------------------------------------')
    MWLOG.info('MW uninstaller start: ' + time.ctime())
    MWLOG.info('---------------------------------------------')

    if len(sys.argv) > 4:
        MWLOG.info('Removing installation folder')
        remove_folders([INST_FOLDER]) # add folders to list if there are more then one in the future
        MWLOG.info('Removing tool bar')
        remove_toolbar(RHINO_FOLDER + TOOLBAR_FILE)
        MWLOG.info('Removing config file')
        remove_conf_file(RHINO_FOLDER)
        MWLOG.info('Removing registry entry')
        remove_reg_key()
        MWLOG.info('Removing start menu entry')
        remove_startmenu_shortcut()
        MWLOG.info('DONE')
    else:
        print 'Folder destinations are missing.'
        MWLOG.error('Folder destinations are missing. Add sys arguments.')

    # move logfile
    try:
        shutil.move(r'MWLOG_UAC.log', SETUP_FOLDER + r'\MWLOG_UAC.log')
    except:
        pass

logging.basicConfig(filename='MWLOG_UAC.log', level=logging.DEBUG)
MWLOG = logging.getLogger('MWSETUP_UAC')

# UAC will start from here
if SETUP_ROUTINE == 'install':
    install()
else:
    uninstall()
