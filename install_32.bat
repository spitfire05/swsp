@echo off
C:\Windows\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe -codebase swsp.dll
if errorlevel 1 goto fail
cscript MessageBox.vbs "Installed for 32 bit SolidWorks. f you change the location of this addin, please re-run the install script."
exit
:fail
cscript MessageBox.vbs "Installation failed!"
