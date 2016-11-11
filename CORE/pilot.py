import os, sys
import win32com.shell.shell as shell
import threading
import shutil


class pilot(threading.Thread):
    def __init__(self, rhino_folder, installation_folder):
        self.rhino_folder = rhino_folder
        self.installation_folder = installation_folder
        self.package_folder = r'../bin/package/'
        self.package_file = 'MWAdditive.zip'
        self.config_file = 'MW_printer_global_parameter.py'
        self.lib_folder = r'Plug-ins\IronPython\Lib\\'
        self.toolbar_folder = r'Plug-ins\Toolbars\\'
        self.toolbar_file = r'MW3DPrint_TB.rui'
        threading.Thread.__init__(self)

    # ------------------------------------------------------------------------------------------------------------------
    def run(self):

        inst_folder_split = self.installation_folder.split('\\')
        print inst_folder_split

        stat = self.install()
        print 'DONE'
        print 'Starting Rhino'

        cmd = '""%s" "%s""' % (self.rhino_folder + r'system\Rhino.exe ',
                               self.rhino_folder + self.toolbar_folder + self.toolbar_file)

        os.system(cmd)

    # ------------------------------------------------------------------------------------------------------------------
    def install(self):
        ASADMIN = 'asadmin'
        script = os.path.abspath('..\\CORE\\UAC_tools.py')
        params = ' '.join([script] +
                          [self.installation_folder.replace(' ', '%20')] +
                          [self.rhino_folder.replace(' ', '%20')] +
                          [self.package_folder.replace(' ', '%20') + self.package_file] +
                          [ASADMIN])

        stat = shell.ShellExecuteEx(lpVerb='runas',
                                    lpFile=sys.executable,
                                    lpParameters=params
                                    )

        return stat
