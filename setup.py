import wx
from UI import installShield


if __name__ == "__main__":
    app = wx.App(False)
    IS = installShield.installShield()
    app.MainLoop()
