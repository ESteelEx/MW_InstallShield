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
        script = os.path.abspath(sys.argv[0])
        params = ' '.join([script] + sys.argv[1:] + [ASADMIN])
        shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)

        self.installation_folder = 'C:\\MWAdditive22\\'

        if not os.path.isdir(self.installation_folder):
            os.mkdir(self.installation_folder)

        #pattern = '*.zip'
        #for root, dirs, files in os.walk(self.package_folder):
        #    for filename in fnmatch.filter(files, pattern):
        #        zipfile.ZipFile(os.path.join(root, filename)).extractall(
        #            os.path.join(self.installation_folder, os.path.splitext(filename)[0]))

        fh = open(self.package_folder + self.package_file, 'rb')
        z = zipfile.ZipFile(fh)
        for name in z.namelist():
            try:
                print name
                if str(name)[-1] == '/':
                    try:

                        folders = self.installation_folder.split('/')
                        folder = ''
                        for folder in folders:
                            folder_str = folder_str + folder
                            os.mkdir(self.installation_folder + folder_str)

                    except:
                        print 'Could not create folder'

                else:
                    outfile = open(self.installation_folder + name, 'wb')
                    outfile.write(z.read(name))
                    outfile.close()
            except:
                print name + ' failed'

        fh.close()
