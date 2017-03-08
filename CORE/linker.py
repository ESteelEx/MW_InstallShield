import os
import win32com.client


shortcut_group_name = "My Start Menu Shortcuts"
shortcut_name = "Visit Microsofts website"
shortcut_target = "http://www.microsoft.com"

sh = win32com.client.Dispatch("WScript.Shell")
p = sh.SpecialFolders("AllUsersPrograms")
assert( os.path.isdir(p) )
p= os.path.join( p, shortcut_group_name)
if ( not os.path.isdir(p) ) :
    os.makedirs(p)

lnk = sh.CreateShortcut( os.path.join( p, shortcut_name + ".lnk" ))
lnk.TargetPath = shortcut_target
lnk.Save()