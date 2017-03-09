@echo off
cd CORE
echo ...
echo COMPILING UAC MODULE
pyinstaller --clean --win-private-assemblies UAC_EXECUTER.spec
cd ..
echo DONE