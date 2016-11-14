import wx, logging
from UI import installShield


if __name__ == "__main__":
    app = wx.App(False)
    logging.basicConfig(filename='MWLOG.log', level=logging.DEBUG)
    MWLOG = logging.getLogger('MWSETUP')
    try:
        IS = installShield.installShield()
    except Exception:
        MWLOG.exception('MAIN')

    app.MainLoop()
