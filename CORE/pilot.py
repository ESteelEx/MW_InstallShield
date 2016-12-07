import os, sys
import win32com.shell.shell as shell
import threading
import psutil
import logging


class pilot(threading.Thread):
    def __init__(self, rhino_folder, installation_folder):
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
        threading.Thread.__init__(self)

    # ------------------------------------------------------------------------------------------------------------------
    def run(self):

        self.MWLOG.info('STARTING INSTALLATION RUN')

        running_inst = 0
        for proc in psutil.process_iter():
            try:
                if proc.name() == 'Rhino.exe':
                    self.MWLOG.info('AN INSTANCE OF RHINO IS RUNNING')
                    print 'Rhino is running. Please exit instance.'
                    try:
                        print 'Killing process '
                        proc.kill()
                        print 'Rhino killed.'
                        self.MWLOG.info('I KILLED A RHINO')
                    except:
                        self.MWLOG.info('COULD NOT KILL RHINO')
                        print 'Wasnt successful. End manually.'
                        print 'Retry!'
                        return

            except:
                self.MWLOG.exception('THREAD')

        stat = self.install()

        print 'Installation done.'

        # print 'Starting Rhino'

        cmd = '""%s" "%s""' % (self.rhino_folder + r'system\Rhino.exe ',
                               self.rhino_folder + self.toolbar_folder + self.toolbar_file)

        # os.system(cmd)

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

        self.MWLOG.info('Workingdirectory: ' + str(exeFolder))

        print exeFolder
        print executable

        params = ' '.join([self.installation_folder.replace(' ', '%20')] +
                          [self.rhino_folder.replace(' ', '%20')] +
                          [exeFolder + self.package_folder.replace(' ', '%20') + self.package_file] +
                          [self.setup_folder.replace(' ', '%20')] +
                          [ASADMIN])

        self.MWLOG.info(params)

        if os.path.isfile(exeFolder + executable):
            self.MWLOG.info('UAC MODULE FOUND. LET MICROSOFT DECIDE WHAT THE HELL WILL HAPPEN.')

        try:
            dict = shell.ShellExecuteEx(lpVerb='runas',
                                        lpFile=exeFolder + executable,
                                        lpParameters=params
                                        )

            hh = dict['hProcess']
            ret = win32event.WaitForSingleObject(hh, -1)

            self.MWLOG.info(str(ret))
            self.MWLOG.info('UAC FINISHED')

            stat = 1

        except Exception:
            stat = -1
            self.MWLOG.exception('SHELL EXECUTE')

        return stat
