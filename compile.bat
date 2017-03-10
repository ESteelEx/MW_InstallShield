@echo off
cd CORE
echo ...
echo COMPILING UAC MODULE
pyinstaller --clean --win-private-assemblies UAC_EXECUTER.spec
echo ...
echo COMPILING SETUP
cd ..
pyinstaller --clean --win-private-assemblies setup.spec
echo DONE