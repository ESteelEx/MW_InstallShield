@echo off
cd CORE
echo ...
echo COMPILING UAC MODULE
pyinstaller UAC_EXECUTER.spec
echo ...
echo COMPILING SETUP
cd ..
pyinstaller setup.spec
echo DONE