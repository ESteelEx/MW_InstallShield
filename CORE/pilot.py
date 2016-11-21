import os, sys
import win32com.shell.shell as shell
import threading
import psutil
import logging
import win32event
import subprocess as sp




class pilot(threading.Thread):
    def __init__(self, rhino_folder, installation_folder):
        self.rhino_folder = rhino_folder
        self.installation_folder = installation_folder
        self.package_folder = r'bin\package\\'
        self.package_file = 'MWAdditive.zip'
        self.config_file = 'MW_printer_global_parameter.py'
        self.lib_folder = r'Plug-ins\IronPython\Lib\\'
        self.toolbar_folder = r'Plug-ins\Toolbars\\'
        self.toolbar_file = r'MW3DPrint_TB.rui'
        self.MWLOG = logging.getLogger('MWSETUP')
        threading.Thread.__init__(self)

    # ------------------------------------------------------------------------------------------------------------------
    def run(self):

        running_inst = 0
        for proc in psutil.process_iter():
            try:
                if proc.name() == 'Rhino.exe':
                    print 'Rhino is running. Please exit instance.'
                    try:
                        print 'Killing process '
                        proc.kill()
                        print 'Rhino killed.'
                    except:
                        print 'Wasnt successful. End manually.'
                        print 'Retry!'
                        return

            except:
                self.MWLOG.exception('THREAD')

        self.MWLOG.info('STARTING INSTALLATION RUN')
        stat = self.install()

        print 'Installation done.'

        # print 'Starting Rhino'

        cmd = '""%s" "%s""' % (self.rhino_folder + r'system\Rhino.exe ',
                               self.rhino_folder + self.toolbar_folder + self.toolbar_file)

        # os.system(cmd)

    # ------------------------------------------------------------------------------------------------------------------
    def install(self):
        ASADMIN = 'asadmin'
        script = os.path.abspath(r'CORE\UAC_tools.py')
        #params = ' '.join([script] +
        #                  [self.installation_folder.replace(' ', '%20')] +
        #                  [self.rhino_folder.replace(' ', '%20')] +
        #                  [self.package_folder.replace(' ', '%20') + self.package_file] +
        #                  [ASADMIN])

        try:
            #dict = shell.ShellExecuteEx(lpVerb='runas',
            #                            lpFile=sys.executable,
            #                            lpParameters=params
            #                            )


            #hh = dict['hProcess']
            #print hh
            #ret = win32event.WaitForSingleObject(hh, -1)
            #   print ret

            params = ' '.join([script] +
                              [self.installation_folder.replace(' ', '%20')] +
                              [self.rhino_folder.replace(' ', '%20')] +
                              [self.package_folder.replace(' ', '%20') + self.package_file])

            prog = sp.Popen(['runas', params, '/noprofile', '/user:Administrator', 'NeedsAdminPrivilege.exe'], stdin=sp.PIPE)
            # prog.stdin.write('password')
            prog.communicate()

            stat = 1

        except Exception:
            stat = -1
            self.MWLOG.exception('SHELL EXECUTE')

        return stat
