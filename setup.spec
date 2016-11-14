import shutil, psutil

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
                datas=[('bin', 'bin')],
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
    debug=True,
    strip=None,
    upx=True,
    console=True,
    icon='bin\\images\\install.ico')

if os.path.isfile('setup.exe'):
    os.remove('setup.exe')
    shutil.copy('dist\\setup.exe', 'setup.exe')
else:
    shutil.copy('dist\\setup.exe', 'setup.exe')

