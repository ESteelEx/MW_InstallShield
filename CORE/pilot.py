import shutil, os, zipfile, sys
import win32com.shell.shell as shell
import zipfile, fnmatch
import threading

class pilot(threading.Thread):
    def __init__(self, rhino_folder, installation_folder):
        self.rhino_folder = rhino_folder
        self.installation_folder = installation_folder
        self.package_folder = r'../bin/package/'
        self.package_file = 'MWAdditive.zip'
        threading.Thread.__init__(self)

    # ------------------------------------------------------------------------------------------------------------------
    def run(self):
        self.extract()

    # ------------------------------------------------------------------------------------------------------------------
    def extract(self):

        ASADMIN = 'asadmin'
        # script = os.path.abspath(sys.argv[0])
        # params = ' '.join([script] + sys.argv[1:] + [ASADMIN])
        script = os.path.abspath('..\\CORE\\UAC_tools.py')
        params = ' '.join([script] + [ASADMIN])
        stat = shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)

        fh = open(self.package_folder + self.package_file, 'rb')
        z = zipfile.ZipFile(fh)
        for name in z.namelist():
            try:
                print name
                if str(name)[-1] == '/':

                        folders = name.split('/')
                        folder_str = ''
                        for folder in folders:
                            folder_str = folder_str + folder + '\\'
                            print 'Creating: ' + self.installation_folder + folder_str
                            try:
                                os.mkdir(self.installation_folder + folder_str)
                            except:
                                pass

                else:
                    outfile = open(self.installation_folder + name, 'wb')
                    outfile.write(z.read(name))
                    outfile.close()
            except:
                print name + ' failed'

        fh.close()
