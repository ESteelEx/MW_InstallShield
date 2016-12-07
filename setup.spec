import shutil, psutil, os

for proc in psutil.process_iter():
    try:
        if proc.name() == 'setup.exe':
            try:
                proc.kill()
            except:
                pass
    except:
        pass

# -*- mode: python -*-
a = Analysis(['setup.py'],
                datas=[('bin', 'bin'), ('CORE', 'CORE'), ('UI', 'UI')],
                pathex=['.'],
                hiddenimports=[],
                hookspath=None,
                runtime_hooks=None)

# a.datas.append(('cacert.pem', 'cacert.pem', 'DATA'))

for d in a.datas:
    if 'pyconfig' in d[0]:
        a.datas.remove(d)
        break

pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name=os.path.join('dist', 'setup.exe'),
    debug=False,
    strip=None,
    upx=True,
    console=False,
    icon='bin\\images\\install.ico')

if os.path.isfile('MW3DPrintingForRhino.exe'):
    os.remove('MW3DPrintingForRhino.exe')
    shutil.copy('dist\\setup.exe', 'MW3DPrintingForRhino.exe')
else:
    shutil.copy('dist\\setup.exe', 'MW3DPrintingForRhino.exe')

