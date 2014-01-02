@echo off
C:\Windows\Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe -codebase swsp.dll
if errorlevel 1 goto fail
cscript MessageBox.vbs "Installed for 64 bit SolidWorks."
exit
:fail
cscript MessageBox.vbs "Installation failed!"
