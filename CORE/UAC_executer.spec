import shutil, psutil

for proc in psutil.process_iter():
    try:
        if proc.name() == 'UAC_EXECUTER.exe':
            try:
                proc.kill()
            except:
                pass
    except:
        pass

# -*- mode: python -*-
a = Analysis(['UAC_tools.py'],
                datas=[],
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
    name=os.path.join('dist', 'UAC_EXECUTER.exe'),
    debug=True,
    strip=None,
    upx=True,
    console=True)

if os.path.isfile('setup.exe'):
    os.remove('setup.exe')
    shutil.copy('dist\\UAC_EXECUTER.exe', 'UAC_EXECUTER.exe')
else:
    shutil.copy('dist\\UAC_EXECUTER.exe', 'UAC_EXECUTER.exe')