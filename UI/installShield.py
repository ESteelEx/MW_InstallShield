import os, wx, sys
import UI_settings as UI
import _winreg
import win32event, win32con, win32process
import psutil, logging
from CORE import pilot
import win32com.shell.shell as shell
from win32com.shell import shellcon


# ----------------------------------------------------------------------------------------------------------------------
def qstDialog(question, caption,):
    dlg = wx.MessageDialog(None,
                           question,
                           caption,
                           wx.YES_NO | wx.ICON_QUESTION)
    result = dlg.ShowModal()
    dlg.Destroy()

    return result

class installShield(wx.Dialog):
    def __init__(self):

        self.MWLOG = logging.getLogger('MWSETUP')
        self.prev_installation = False

        try:
            base_path = sys._MEIPASS + '\\'
            print 'Start from EXE ... Working in ->'
            print base_path
        except Exception:
            print 'Start from script ... Working in ->'
            base_path = os.path.abspath(".") + '\\'
            print base_path

        self.root_reg_key = r'Software\McNeel\Rhinoceros\\'
        self.MW_reg_key = r'Software\MW\3DPrinting\\'

        self.postfix_installation_folder = r'\MW3DPrinting\\'
        self.icon_main = base_path + r'bin\images\install.ico'

        self.toolbar_folder = r'Plug-ins\Toolbars\\'
        self.toolbar_file = r'MW3DPrint_TB.rui'

        self.editbox = []

        wx.Dialog.__init__(self, None, title='MW Installer - 3D Printing for Rhino', size=UI.WMAIN['size'],
                           style=wx.SYSTEM_MENU |
                                 wx.CAPTION |
                                 wx.CLOSE_BOX |
                                 wx.TAB_TRAVERSAL)  # | wx.RESIZE_BORDER)  # | wx.TRANSPARENT_WINDOW)

        self.win_style = self.GetWindowStyle()
        self.SetWindowStyle(self.win_style | wx.STAY_ON_TOP)

        icon = wx.IconFromBitmap(wx.Bitmap(self.icon_main))
        self.SetIcon(icon)

        # self.png = wx.StaticBitmap(self, -1, wx.Bitmap(r"bin\images\down.png", wx.BITMAP_TYPE_ANY))

        # self.png.Pos

        self.SetSizeWH(UI.WMAIN['size'][0], UI.WMAIN['size'][1])

        self.SetBackgroundColour(wx.Colour(UI.WCOLOR['BG'][0],
                                           UI.WCOLOR['BG'][1],
                                           UI.WCOLOR['BG'][2]))

        atable = wx.AcceleratorTable([(wx.ACCEL_NORMAL, wx.WXK_ESCAPE, wx.ID_EXIT)])
        self.SetAcceleratorTable(atable)
        wx.EVT_MENU(self, wx.ID_EXIT, self.close_IS)

        self.Bind(wx.EVT_CLOSE, self.close_IS)

        self.Center()
        self.Show()

        self.current_y_pxpos_elem = 0

        # TEXT HEADER
        # _______________________________________________________________________________

        label = 'Please enter information below to start MW 3D Printer installation.'

        text = wx.StaticText(self,
                             label=label,
                             pos=(UI.THEADERSTART['pos'][0],
                                  UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem))

        text.SetForegroundColour(UI.TCOLOR['FG'])  # set text color

        self.current_y_pxpos_elem += 40

        # TEXT RHINO
        # _______________________________________________________________________________

        label = 'Rhino installation folder:'

        text = wx.StaticText(self,
                             label=label,
                             pos=(UI.THEADERSTART['pos'][0],
                                  UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem))

        text.SetForegroundColour(UI.ECOLOR2['FG'])  # set text color

        self.current_y_pxpos_elem += 15

        # EDITBOX RHINO
        # _______________________________________________________________________________

        self.editbox.append( (wx.TextCtrl(self,
                               name='path_Rhino',
                               pos=(UI.THEADERSTART['pos'][0], UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem),
                               size=(UI.WMAIN['size'][0]-45, UI.EBOX['size'][1]),
                               style=wx.TE_PROCESS_ENTER)) )

        self.editbox[0].SetForegroundColour(UI.ECOLOR2['FG'])  # set color
        self.editbox[0].SetBackgroundColour(UI.WCOLOR['BG'])  # set color
        self.editbox[0].Bind(wx.EVT_KEY_DOWN, self.installOK)

        try:
            print 'Searching for Rhino installation ...'
            key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, self.root_reg_key)
        except:
            key = None
            message = 'There is no Rhino installed. Visit http://www.rhino3d.com/download/ to get your version.'
            print message
            return

        if key is not None:
            i = 0
            while 1:
                try:
                    subkey =_winreg.EnumKey(key, i)
                    if subkey.find('x') != -1:
                        print 'Rhino version ' + subkey.split('x')[0] + ' - ' + subkey.split('x')[1] + ' bit found.'

                        label = 'Rhino version ' + subkey.split('x')[0] + ' - ' + subkey.split('x')[1] + ' bit - found'

                        text = wx.StaticText(self,
                                             label=label,
                                             pos=(UI.THEADERSTART['pos'][0],
                                                  UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem))

                        text.SetForegroundColour(UI.PARAMCOLOR['FG'])  # set text color

                        self.current_y_pxpos_elem += 15

                        self.RhinoFound = True

                        break

                    i += 1
                except:
                    break

            key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, self.root_reg_key + subkey + "\\Install")
            reg_value = _winreg.QueryValueEx(key, "InstallPath")[0]

            self.reg_rhino_folder = reg_value

            self.editbox[0].SetValue(reg_value)
            self.editbox[0].MoveXY(UI.THEADERSTART['pos'][0], UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem)

        else:
            self.RhinoFound = False

            label = 'No Rhino installation found.\n' \
                    'Visit www.Rhino3d.com to download the latest version.'

            text = wx.StaticText(self,
                                 label=label,
                                 pos=(UI.THEADERSTART['pos'][0],
                                      UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem))

            text.SetForegroundColour(UI.TERROR['FG'])  # set text color

            self.current_y_pxpos_elem += 15
            self.editbox[0].Hide()

        self.current_y_pxpos_elem += 30

        # TEXT INSTALL CORE
        # _______________________________________________________________________________

        label = 'Enter installation folder. Do not install into ProgramFiles folder:'

        text = wx.StaticText(self,
                             label=label,
                             pos=(UI.THEADERSTART['pos'][0],
                                  UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem))

        text.SetForegroundColour(UI.ECOLOR2['FG'])  # set text color

        self.current_y_pxpos_elem += 15

        # EDITBOX CORE PATH
        # _______________________________________________________________________________

        self.editbox.append((wx.TextCtrl(self,
                               name='path_Rhino',
                               pos=(UI.THEADERSTART['pos'][0], UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem),
                               size=(UI.WMAIN['size'][0]-70, UI.EBOX['size'][1]),
                               style=wx.TE_PROCESS_ENTER)))

        self.editbox[1].SetForegroundColour(UI.ECOLOR2['FG'])  # set color
        self.editbox[1].SetBackgroundColour(UI.WCOLOR['BG'])  # set color
        self.editbox[1].Bind(wx.EVT_KEY_DOWN, self.installOK)

        try:
            key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, self.MW_reg_key)
            core_path = _winreg.QueryValue(key, 'CorePath')
            self.reg_core_path = core_path
            self.config_file = _winreg.QueryValue(key, 'ConfigFile')

            if core_path[-1] == '\\\\':
                core_path = core_path.decode('string_escape')

            label = 'Previous installation found in'
            self.prev_installation = True

            text = wx.StaticText(self,
                                 label=label,
                                 pos=(UI.THEADERSTART['pos'][0],
                                      UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem))

            text.SetForegroundColour(UI.PARAMCOLOR['FG'])  # set text color

            self.current_y_pxpos_elem += 15

            self.editbox[1].SetValue(core_path)
            self.editbox[1].MoveXY(UI.THEADERSTART['pos'][0], UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem)
        except Exception as e:
            self.MWLOG.exception('REGKEY')
            self.editbox[1].SetValue(os.environ['PROGRAMFILES'][0:2] + self.postfix_installation_folder.decode('string_escape'))


        # BUTTON CHOOSE FOLDER
        # _______________________________________________________________________________
        self.button_choose_folder = wx.Button(self,
                                              label='...',
                                              size=[20, 20],
                                              pos=[455, UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem])

        self.button_choose_folder.SetForegroundColour((250, 250, 250))
        self.button_choose_folder.SetBackgroundColour((80, 80, 80))
        self.button_choose_folder.Bind(wx.EVT_BUTTON, self.choose_folder)

        self.current_y_pxpos_elem += 30

        # CHECKBOX
        # _______________________________________________________________________________
        # CheckBox(parent, id=ID_ANY, label="", pos=DefaultPosition,
        #          size=DefaultSize, style=0, validator=DefaultValidator,
        #          name=CheckBoxNameStr)

        # checkbox = wx.CheckBox(self,
        #                       label='Install Rhino tool bar',
        #                       pos=(UI.THEADERSTART['pos'][0], UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem))

        # checkbox.SetForegroundColour(UI.ECOLOR2['FG'])
        # checkbox.SetValue(1)


        # BUTTON (INSTALL)
        # _______________________________________________________________________________

        label = 'INSTALL'

        pos = (UI.WMAIN['size'][0] - 70, UI.WMAIN['size'][1] - 50)

        self.button_install = wx.StaticText(self,
                                            label=label,
                                            pos=pos)

        self.button_install.SetForegroundColour(UI.ECOLOR2['FG'])

        self.button_install.Bind(wx.EVT_MOTION, self.installOver)
        self.button_install.Bind(wx.EVT_LEFT_DOWN, self.installDOWN)

        if self.RhinoFound:
            self.button_install.Bind(wx.EVT_LEFT_UP, self.installOK)
        else:
            self.button_install.SetLabel('EXIT')
            self.button_install.Bind(wx.EVT_LEFT_UP, self.close_IS)


        # BUTTON (UNINSTALL)
        # __________________________________________________________________

        if self.prev_installation:

            label = 'UNINSTALL'

            pos = (UI.WMAIN['size'][0] - 480, UI.WMAIN['size'][1] - 50)

            self.button_uninstall = wx.StaticText(self,
                                                  label=label,
                                                  pos=pos)

            self.button_uninstall.SetForegroundColour(UI.ECOLOR2['FG'])

            self.button_uninstall.Bind(wx.EVT_MOTION, self.uninstallOver)
            self.button_uninstall.Bind(wx.EVT_LEFT_DOWN, self.uninstallDOWN)

            self.button_uninstall.Bind(wx.EVT_LEFT_UP, self.uninstallOK)

    # ------------------------------------------------------------------------------------------------------------------
    def choose_folder(self, event):
        # remove on top
        self.SetWindowStyle(self.win_style)

        dlg = wx.DirDialog(None, 'Choose input directory', style=wx.DD_DEFAULT_STYLE)

        # self is a frame
        style = self.GetWindowStyle()
        # stay on top
        self.SetWindowStyle(style | wx.STAY_ON_TOP)
        # normal behaviour again
        self.SetWindowStyle(style)

        if dlg.ShowModal() == wx.ID_OK:
            self.editbox[1].SetValue(dlg.GetPath())

        dlg.Destroy()

        # add on top
        self.SetWindowStyle(self.win_style | wx.STAY_ON_TOP)

    # ------------------------------------------------------------------------------------------------------------------
    def installOver(self, event):
        pass

    # ------------------------------------------------------------------------------------------------------------------
    def installDOWN(self, event):
        if self.RhinoFound:
            self.button_install.SetLabel('INSTALL')
        else:
            self.button_install.SetLabel('EXIT')
        self.button_install.SetForegroundColour(UI.SWAPPING['DOWN'])



    # ------------------------------------------------------------------------------------------------------------------
    def installOK(self, event):
        self.button_install.SetLabel('INSTALL')
        self.button_install.SetForegroundColour(UI.ECOLOR2['FG'])

        try:
            keycode = event.GetKeyCode()
            if keycode == wx.WXK_RETURN or keycode == wx.WXK_NUMPAD_ENTER:
                print "you pressed ENTER!"
            else:
                event.Skip()
                return
        except:
            pass

        self.rhino_folder = self.editbox[0].GetValue()
        self.inst_folder = self.editbox[1].GetValue()

        if self.rhino_folder.strip()[-1] != '\\':
            self.rhino_folder = self.rhino_folder + '\\'

        if self.inst_folder.strip()[-1] != '\\':
            self.inst_folder = self.inst_folder + '\\'

        print self.rhino_folder
        print self.inst_folder

        for proc in psutil.process_iter():
            try:
                if proc.name() == 'Rhino.exe':
                    self.MWLOG.info('AN INSTANCE OF RHINO IS RUNNING')
                    print 'Rhino is running. Please exit instance.'

                    answer = qstDialog('You want to close Rhino instance? All data will be lost.',
                                       'Closing process')

                    if answer == wx.ID_YES:
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
                    else:
                        return

            except Exception as e:
                if len(e.message) > 0:
                    self.MWLOG.error(e.message)

        P = pilot.pilot(self.rhino_folder, self.inst_folder)
        P.start()
        P.join()

        answer = qstDialog('Do you want to start Rhino now to activate tool bar? Otherwise start it from MW folder in start menu the first time.',
                           'Activation')

        if answer == wx.ID_YES:
            cmd = self.rhino_folder + r'system\Rhino.exe'
            params = '"' + self.rhino_folder + self.toolbar_folder + self.toolbar_file + '"'

            showCmd = win32con.SW_HIDE  # SW_SHOWNORMAL

            procInfo = shell.ShellExecuteEx(nShow=showCmd,
                                            fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
                                            lpFile=cmd,
                                            lpParameters=params
                                            )

        # self.Destroy()

    # ------------------------------------------------------------------------------------------------------------------
    def uninstallOver(self, event):
        pass

    # ------------------------------------------------------------------------------------------------------------------
    def uninstallDOWN(self, event):
        self.button_uninstall.SetLabel('UNINSTALL')
        self.button_uninstall.SetForegroundColour(UI.SWAPPING['DOWN'])

    # ------------------------------------------------------------------------------------------------------------------
    def uninstallOK(self, event):
        self.button_uninstall.SetLabel('UNINSTALL')
        self.button_uninstall.SetForegroundColour(UI.ECOLOR3['FG'])

        answer = qstDialog('Are you sure you want to remove MW 3D Printing for Rhino?',
                           'Uninstall')

        if answer == wx.ID_YES:

            print self.reg_rhino_folder
            print self.reg_core_path

            for proc in psutil.process_iter():
                try:
                    if proc.name() == 'Rhino.exe':
                        self.MWLOG.info('AN INSTANCE OF RHINO IS RUNNING')
                        print 'Rhino is running. Please exit instance.'

                        answer = qstDialog('You want to close Rhino instance? All data will be lost.',
                                           'Closing process')

                        if answer == wx.ID_YES:
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
                        else:
                            return

                except Exception as e:
                    if len(e.message) > 0:
                        self.MWLOG.error(e.message)

            P = pilot.pilot(self.reg_rhino_folder, self.reg_core_path, routine='uninstall')
            P.start()
            P.join()

            # self.Destroy()


    # ------------------------------------------------------------------------------------------------------------------
    def close_IS(self, event):
        print 'Canceled by user.'
        self.Destroy()

