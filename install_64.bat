@echo off
C:\Windows\Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe -codebase swsp.dll
if errorlevel 1 goto fail
cscript MessageBox.vbs "Installed for 64 bit SolidWorks. If you change the location of this addin, please re-run the install script."
exit
:fail
cscript MessageBox.vbs "Installation failed!"
