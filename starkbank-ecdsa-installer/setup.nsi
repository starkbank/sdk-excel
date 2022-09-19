Function registerEnvirVar
   ; include for some of the windows messages defines
   !include "winmessages.nsh"
   ; HKLM (all users) defines
   !define env_hklm 'HKLM "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"'
   ; initialize envirnoment variable name and value
   !define varName "StarkBankExcelEcdsa"
   !define varValue $INSTDIR
   ; set variable for local machine
   WriteRegExpandStr ${env_hklm} ${varName} ${varValue}
   ; make sure windows knows about the change
   SendMessage ${HWND_BROADCAST} ${WM_WININICHANGE} 0 "STR:Environment" /TIMEOUT=5000
FunctionEnd

;--------------------------------
; Includes

  !include "MUI2.nsh"
  !include "logiclib.nsh"

;--------------------------------
; Custom defines
  !define NAME "StarkBankExcelEcdsa"
  !define VERSION "2.0.0"
  !define SLUG "${NAME} v${VERSION}"

;--------------------------------
; Generals

  Name "${NAME}"
  OutFile "${NAME} Setup.exe"
  InstallDir "$PROGRAMFILES\${NAME}"
  InstallDirRegKey HKCU "StarkBank\${NAME}" ""
  RequestExecutionLevel admin

;--------------------------------
; UI
  
!define MUI_ICON "assets\ICON.ico"
!define MUI_HEADERIMAGE
!define MUI_WELCOMEFINISHPAGE_BITMAP "assets\welcome.bmp"
!define MUI_ABORTWARNING
!define MUI_WELCOMEPAGE_TITLE "${SLUG} Setup"

;--------------------------------
; Pages
  
; Installer pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "license.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Set UI language
!insertmacro MUI_LANGUAGE "English"

;--------------------------------
; Section - Install App

Section "-hidden app"
  SectionIn RO
  SetOutPath "$INSTDIR"
  File /r "app\*.*" 
  WriteRegStr HKCU "Software\${NAME}" "" $INSTDIR
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  call registerEnvirVar
SectionEnd
