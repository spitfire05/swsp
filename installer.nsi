!include x64.nsh
!include LogicLib.nsh
!include MUI2.nsh

!define PROGNAME "swsp"
!define VERSION "1.0.1"

;--------------------------------

; The name of the installer
Name "SolidWorks Standard Primitives ${VERSION}"

OutFile "${PROGNAME}-install-${VERSION}.exe"

; Registry key to check for directory (so if you install again, it will 
; overwrite the old one automatically)
InstallDirRegKey HKLM "Software\spitfire_swsp" "Install_Dir"

; Request application privileges for Windows Vista
RequestExecutionLevel admin

;--------------------------------

; Functions

Function .onInit
  ${If} ${RunningX64}
    SetRegView 64
    StrCpy $INSTDIR $PROGRAMFILES64\${PROGNAME}
  ${Else}
    StrCpy $INSTDIR $PROGRAMFILES64\${PROGNAME}
  ${EndIf}
FunctionEnd

;--------------------------------

;Interface Settings

  !define MUI_ABORTWARNING

;--------------------------------
;Pages

  !insertmacro MUI_PAGE_DIRECTORY
  !insertmacro MUI_PAGE_INSTFILES
  
  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES
  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "English"

;--------------------------------

; The stuff to install
Section "Install"

  SectionIn RO
  
  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ; Put file there
  File "bin\Release\swsp.dll"
  File "bin\Release\SolidWorks.Interop.sldworks.dll"
  File "bin\Release\SolidWorks.Interop.swconst.dll"
  File "bin\Release\SolidWorks.Interop.swpublished.dll"
  File "readme.txt"
  File "changelog.txt"
  CreateDirectory $INSTDIR\icons
  File /oname=icons\icons_16.png "bin\Release\icons\icons_16.png"
  
  ; Register with SolidWorks
  ${If} ${RunningX64}
    ExecWait '"$WINDIR\Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe" "-codebase" "$INSTDIR\swsp.dll"' $0
  ${Else}
    ExecWait '"$WINDIR\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe" "-codebase" "$INSTDIR\swsp.dll"'
  ${EndIf}
  IfErrors 0 +2
    Abort "Failed to Register addin with SolidWorks. Make sure you have .net Framework v3.5 installed."
  ClearErrors
  
  
  ; Write the installation path into the registry
  WriteRegStr HKLM SOFTWARE\spitfire_swsp "Install_Dir" "$INSTDIR"
  
  ; Write the uninstall keys for Windows
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\spitfire_swsp" "DisplayName" "SolidWorks Standard Primitives"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\spitfire_swsp" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\spitfire_swsp" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\spitfire_swsp" "NoRepair" 1
  WriteUninstaller "uninstall.exe"
  
SectionEnd

;--------------------------------

; Uninstaller

Section "Uninstall"
  
  ; Remove registry keys
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\spitfire_swsp"
  DeleteRegKey HKLM SOFTWARE\spitfire_swsp
  
  ; Un-register with SolidWorks
  ${If} ${RunningX64}
    ExecWait '"$WINDIR\Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe" "-unregister" "$INSTDIR\swsp.dll"' $0
  ${Else}
    ExecWait '"$WINDIR\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe" "-unregister" "$INSTDIR\swsp.dll"'
  ${EndIf}
  
  ; Remove files and uninstaller
  Delete "$INSTDIR\swsp.dll"
  Delete "$INSTDIR\SolidWorks.Interop.sldworks.dll"
  Delete "$INSTDIR\SolidWorks.Interop.swconst.dll"
  Delete "$INSTDIR\SolidWorks.Interop.swpublished.dll"
  Delete "$INSTDIR\readme.txt"
  Delete "$INSTDIR\changelog.txt"
  Delete "$INSTDIR\icons\icons_16.png"
  Delete $INSTDIR\uninstall.exe

  ; Remove directories used
  RMDir "$INSTDIR\icons"
  RMDir "$INSTDIR"

SectionEnd
