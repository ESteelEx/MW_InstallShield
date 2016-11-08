import os, wx
from CORE import pilot
import UI_settings as UI
import _winreg


class installShield(wx.Dialog):
    def __init__(self):

        self.root_reg_key = 'Software\\McNeel\\Rhinoceros'
        self.postfix_installation_folder = u'\MW3DPrinting\\'
        self.icon_main = r'..\bin\images\install.ico'
        self.editbox = []

        wx.Dialog.__init__(self, None, title='MW Installer - 3D Printing for Rhino', size=UI.WMAIN['size'],
                           style=wx.SYSTEM_MENU |
                                 wx.CAPTION |
                                 wx.CLOSE_BOX |
                                 wx.TAB_TRAVERSAL |
                                 wx.STAY_ON_TOP )  # | wx.RESIZE_BORDER)  # | wx.TRANSPARENT_WINDOW)

        icon = wx.IconFromBitmap(wx.Bitmap(self.icon_main))
        self.SetIcon(icon)

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


        label = 'Please enter information below to start MW 3D Printer installation.\n' \
                'You need a 32 or 64 bit version of MC Neels Rhino installed.\n' \
                'Visit www.Rhino3d.com to download the latest version.'

        text = wx.StaticText(self,
                             label=label,
                             pos=(UI.THEADERSTART['pos'][0],
                                  UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem))

        text.SetForegroundColour(UI.TCOLOR['FG'])  # set text color

        self.current_y_pxpos_elem += 60

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

        try:
            key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, self.root_reg_key)
        except:
            key = None

        if key is not None:
            i = 0
            while 1:
                try:
                    subkey =_winreg.EnumKey(key, i)
                    if subkey.find('x') != -1:
                        print 'Version found'
                        print subkey.split('x')[1]

                        label = 'Rhino version ' + subkey.split('x')[0] + ' - ' + subkey.split('x')[1] + ' bit - found'

                        text = wx.StaticText(self,
                                             label=label,
                                             pos=(UI.THEADERSTART['pos'][0],
                                                  UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem))

                        text.SetForegroundColour(UI.PARAMCOLOR['FG'])  # set text color

                        self.current_y_pxpos_elem += 15

                        break

                    i += 1
                except:
                    break

            key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, "Software\\McNeel\\Rhinoceros\\" + subkey + "\\Install")
            reg_value = _winreg.QueryValueEx(key, "InstallPath")[0]

            self.editbox[0].SetValue(reg_value)
            self.editbox[0].MoveXY(UI.THEADERSTART['pos'][0], UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem)

        else:
            self.editbox[0].SetValue('No Rhino installation found.')
            self.editbox[0].Enable(False)

        self.current_y_pxpos_elem += 30

        # TEXT INSTALL CORE
        # _______________________________________________________________________________

        label = 'Enter installation folder:'

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
                               size=(UI.WMAIN['size'][0]-45, UI.EBOX['size'][1]),
                               style=wx.TE_PROCESS_ENTER)))

        self.editbox[1].SetForegroundColour(UI.ECOLOR2['FG'])  # set color
        self.editbox[1].SetBackgroundColour(UI.WCOLOR['BG'])  # set color

        self.editbox[1].SetValue(os.environ['PROGRAMFILES'] + self.postfix_installation_folder)

        self.current_y_pxpos_elem += 30

        # CHECKBOX
        # _______________________________________________________________________________
        # CheckBox(parent, id=ID_ANY, label="", pos=DefaultPosition,
        #          size=DefaultSize, style=0, validator=DefaultValidator,
        #          name=CheckBoxNameStr)

        checkbox = wx.CheckBox(self,
                               label='Install Rhino tool bar',
                               pos=(UI.THEADERSTART['pos'][0], UI.THEADERSTART['pos'][1] + self.current_y_pxpos_elem))

        checkbox.SetForegroundColour(UI.ECOLOR2['FG'])
        checkbox.SetValue(1)

        # BUTTON
        # _______________________________________________________________________________

        label = 'INSTALL'

        pos = (UI.WMAIN['size'][0]-70, UI.WMAIN['size'][1]-50)

        self.button_install = wx.StaticText(self,
                                            label=label,
                                            pos=pos)

        self.button_install.SetForegroundColour(UI.ECOLOR2['FG'])
        self.button_install.Bind(wx.EVT_MOTION, self.installOver)
        self.button_install.Bind(wx.EVT_LEFT_DOWN, self.installDOWN)
        self.button_install.Bind(wx.EVT_LEFT_UP, self.installOK)

    # ------------------------------------------------------------------------------------------------------------------
    def installOver(self, event):
        pass

    # ------------------------------------------------------------------------------------------------------------------
    def installDOWN(self, event):
        self.button_install.SetLabel('INSTALL')
        self.button_install.SetForegroundColour(UI.SWAPPING['DOWN'])

    # ------------------------------------------------------------------------------------------------------------------
    def installOK(self, event):
        self.button_install.SetLabel('INSTALL')
        self.button_install.SetForegroundColour(UI.SWAPPING['UP'])

        print self.editbox[0].GetValue()
        print self.editbox[1].GetValue()

        P = pilot.pilot(self.editbox[0].GetValue(), self.editbox[1].GetValue())
        P.start()

    # ------------------------------------------------------------------------------------------------------------------
    def close_IS(self, event):
        self.Destroy()


if __name__ == "__main__":
    app = wx.App(False)
    IS = installShield()
    app.MainLoop()
