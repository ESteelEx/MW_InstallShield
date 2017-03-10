import os, sys
import logging, threading
import win32event, win32con, win32process

import win32com.shell.shell as shell
from win32com.shell import shellcon


class pilot(threading.Thread):
    def __init__(self, rhino_folder, installation_folder, routine='install'):
        self.routine = routine
        self.rhino_folder = rhino_folder
        self.installation_folder = installation_folder
        self.setup_folder = os.getcwd()
        self.package_folder = r'bin\package\\'
        self.package_file = r'MWAdditive.zip'
        self.config_file = r'MW_printer_global_parameter.py'
        self.lib_folder = r'Plug-ins\IronPython\Lib\\'
        self.toolbar_folder = r'Plug-ins\Toolbars\\'
        self.toolbar_file = r'MW3DPrint_TB.rui'
        self.MWLOG = logging.getLogger('MWSETUP')
        self.stat = 0
        threading.Thread.__init__(self)

    # ------------------------------------------------------------------------------------------------------------------
    def run(self):

        self.MWLOG.info('STARTING INSTALLATION RUN')

        if self.routine == 'install':
            stat = self.install()
            print 'Installation done.'
        else:
            stat = self.uninstall()
            print 'Uninstall done.'

        self.MWLOG.info('DONE')

    # ------------------------------------------------------------------------------------------------------------------
    def install(self):
        self.MWLOG.info('ENTERING UAC')
        ASADMIN = 'asadmin'
        try:
            exeFolder = sys._MEIPASS
            executable = r'\CORE\UAC_EXECUTER.exe'
        except:
            # exeFolder = os.path.abspath(r'CORE\UAC_EXECUTER.exe')
            exeFolder = os.getcwd()
            executable = r'\CORE\UAC_EXECUTER.exe'

        self.MWLOG.info('WD: ' + str(exeFolder))

        print exeFolder
        print executable

        params = ' '.join([self.installation_folder.replace(' ', '%20')] +
                          [self.rhino_folder.replace(' ', '%20')] +
                          [exeFolder + r'\\' + self.package_folder.replace(' ', '%20') + self.package_file] +
                          [self.setup_folder.replace(' ', '%20')] +
                          ['install'] +
                          [ASADMIN])

        self.MWLOG.info(params)

        if os.path.isfile(exeFolder + executable):
            self.MWLOG.info('UAC MODULE FOUND')

        showCmd = win32con.SW_HIDE  # SW_SHOWNORMAL

        try:
            self.MWLOG.info('UAC EXECUTE')
            procInfo = shell.ShellExecuteEx(nShow=showCmd,
                                            fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
                                            lpVerb='runas',
                                            lpFile=exeFolder + executable,
                                            lpParameters=params
                                            )

            procHandle = procInfo['hProcess']
            obj = win32event.WaitForSingleObject(procHandle, win32event.INFINITE)
            rc = win32process.GetExitCodeProcess(procHandle)

            if rc == 0:
                self.MWLOG.info('UAC STAGE WAS SUCCESSFUL')
            else:
                self.MWLOG.error('EXITCODE OF UAC STAGE: ' + str(rc))

            self.MWLOG.info('UAC FINISHED')

            self.stat = 1

        except Exception:
            self.stat = -1
            self.MWLOG.exception('SHELL EXECUTE')

        return self.stat

    # ------------------------------------------------------------------------------------------------------------------
    def uninstall(self):
        self.MWLOG.info('ENTERING UAC')
        ASADMIN = 'asadmin'
        try:
            exeFolder = sys._MEIPASS
            executable = r'\CORE\UAC_EXECUTER.exe'
        except:
            # exeFolder = os.path.abspath(r'CORE\UAC_EXECUTER.exe')
            exeFolder = os.getcwd()
            executable = r'\CORE\UAC_EXECUTER.exe'

        self.MWLOG.info('WD: ' + str(exeFolder))

        print exeFolder
        print executable

        params = ' '.join([self.installation_folder.replace(' ', '%20')] +
                          [self.rhino_folder.replace(' ', '%20')] +
                          [exeFolder + r'\\' + self.package_folder.replace(' ', '%20') + self.package_file] +
                          [self.setup_folder.replace(' ', '%20')] +
                          ['uninstall'] +
                          [ASADMIN])

        self.MWLOG.info(params)

        if os.path.isfile(exeFolder + executable):
            self.MWLOG.info('UAC MODULE FOUND')

        showCmd = win32con.SW_HIDE  # SW_SHOWNORMAL

        try:
            self.MWLOG.info('UAC EXECUTE')
            procInfo = shell.ShellExecuteEx(nShow=showCmd,
                                            fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
                                            lpVerb='runas',
                                            lpFile=exeFolder + executable,
                                            lpParameters=params
                                            )

            procHandle = procInfo['hProcess']
            obj = win32event.WaitForSingleObject(procHandle, win32event.INFINITE)
            rc = win32process.GetExitCodeProcess(procHandle)

            if rc == 0:
                self.MWLOG.info('UAC STAGE WAS SUCCESSFUL')
            else:
                self.MWLOG.error('EXITCODE OF UAC STAGE: ' + str(rc))

            self.MWLOG.info('UAC FINISHED')

            self.stat = 1

        except Exception:
            self.stat = -1
            self.MWLOG.exception('SHELL EXECUTE')

        return self.stat