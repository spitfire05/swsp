@echo off
C:\Windows\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe -codebase swsp.dll
if errorlevel 1 goto fail
cscript MessageBox.vbs "Installed for 32 bit SolidWorks."
exit
:fail
cscript MessageBox.vbs "Installation failed!"
