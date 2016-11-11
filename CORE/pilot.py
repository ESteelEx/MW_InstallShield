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
        threading.Thread.__init__(self)

    # ------------------------------------------------------------------------------------------------------------------
    def run(self):
        stat = self.install()

        print stat
        print dir(stat['hProcess'].handle)

        print 'DONE'

    # ------------------------------------------------------------------------------------------------------------------
    def install(self):
        ASADMIN = 'asadmin'
        script = os.path.abspath('..\\CORE\\UAC_tools.py')
        params = ' '.join([script] + [self.installation_folder.replace(' ', '%20')] + [self.rhino_folder.replace(' ', '%20')] + [self.package_folder.replace(' ', '%20') + self.package_file] + [ASADMIN])

        stat = shell.ShellExecuteEx(lpVerb='runas',
                                    lpFile=sys.executable,
                                    lpParameters=params
                                    )

        return stat
