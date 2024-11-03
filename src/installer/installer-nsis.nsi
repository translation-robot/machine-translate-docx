!include "MUI2.nsh"  ; Include the Modern UI macro
!include "nsDialogs.nsh"

; Other definitions...

; SetCompressor /SOLID lzma

!define APP_NAME "Machine Translation Robot"
Name "${APP_NAME}"  ; Set the installer title
!define APP_VERSION "0.2"
!define APP_PUBLISHER "Blue Sun"
!define APP_PUBLISHER_URL "https://github.com/translation-robot/machine-translate-docx"
!define DEFAULT_DIR "C:\SMTVRobot"

!define MUI_ICON "C:\SMTVRobot-installer\source_code\img\app.ico"
!define MUI_UNICON "C:\SMTVRobot-installer\source_code\img\app.ico"


!define MUI_WELCOME_TITLE "Welcome to ${APP_NAME} Setup"
!define MUI_WELCOME_TEXT "Welcome to the installation of ${APP_NAME}!"
!define MUI_WELCOME_SUBTITLE "This installer will guide you through the setup process."
!define MUI_WELCOME_FINISH_SUBTITLE "Thank you for installing ${APP_NAME}!"
!define MUI_UNFINISHED_UNINSTALL "1"

;--------------------------------
;defines MUST come before pages to apply to them (in hindsight: duh!)

!define MUI_PAGE_HEADER_TEXT "Machine Translation Robot Compotents:"
!define MUI_PAGE_HEADER_SUBTEXT "Blue Sun September 22nd 2024. Source code on Github. "

!define MUI_WELCOMEPAGE_TITLE "Machine Translation Robot"
!define MUI_WELCOMEPAGE_TEXT "Welcome to Machine Translation Robot Installer. This program translate Word DOCX files using either Deepl or Google translate. It is free of charge and does not require any account. It is possible to use Deepl PRO using account and password. They should be stored manually in configuration.json, contact smtv.bot@gmail.com for any assistance."
;Extra space for the title area
;!insertmacro MUI_WELCOMEPAGE_TITLE_3LINES


; Add the components page here
!insertmacro MUI_PAGE_WELCOME
Page custom InfoPage1 ; Custom Info Page
!insertmacro MUI_PAGE_COMPONENTS  ; Ensure this line is included
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_LANGUAGE "English"

; Define the finish page text
!define MUI_FINISH_TEXT "Installation Complete!"

!define MUI_FINISH_TITLE "Installation Complete!"


; Optionally, you can add an icon
!define MUI_FINISH_ICON "${NSIS_ICON}"

; To customize the finish button text (optional)
!define MUI_FINISH_BUTTON_TEXT "Finish"

; Remove the directory page by not including it
; !insertmacro MUI_PAGE_DIRECTORY

; This section remains unchanged





Section "Core" SEC01
	; Write the uninstall registry entry
	WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayName" "${APP_NAME}"
	WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString" "$INSTDIR\uninstall.exe"


  SectionIn RO
  SetOutPath "$INSTDIR"
  File /r "C:\SMTVRobot-installer\*"
  WriteUninstaller "$INSTDIR\uninstall.exe"
SectionEnd

!define OUTPUT_FILE "installer-nsis.exe"

OutFile "${OUTPUT_FILE}"
InstallDir "${DEFAULT_DIR}"

RequestExecutionLevel user  ; Set the execution level

LangString LANG_ENGLISH ${LANG_ENGLISH} "English"


Section "Desktop Shortcut" SEC02
  CreateShortcut "$DESKTOP\Machine Translation Robot.lnk" "$INSTDIR\bin\machine_translate_gui.exe" "$INSTDIR\bin\app.ico"
SectionEnd

!include "MUI2.nsh"  ; Include the Modern UI macros

OutFile "${APP_NAME} Setup.exe"
InstallDir "${DEFAULT_DIR}"




SectionGroup "Shortcuts" SEC_SHORTCUTS
  Section "Machine Translation Robot" SEC_MTR
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\Machine Translation Robot.lnk"
  SectionEnd

  Section "Graphical Interface - Ask Language" SEC_GI
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe   Graphical Inface - Ask language.lnk"
  SectionEnd

  Section "Show Version" SEC_SV
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe   Show version.lnk"
  SectionEnd

  Section "Split Translation - Any Language" SEC_ST
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe   Split translation - any language.lnk"
  SectionEnd

  Section /o "Machine Translation - Bulgarian - Deepl" SEC_BUL_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Bulgarian - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - Bulgarian - Google" SEC_BUL_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Bulgarian - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Chinese (Simplified) - Deepl" SEC_CHI_SIM_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Chinese (Simplified) - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - Chinese (Traditional) - Google" SEC_CHI_TRA_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Chinese (Traditional) - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Czech - Deepl" SEC_CZE_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Czech - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - Czech - Google" SEC_CZE_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Czech - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - French - Deepl" SEC_FRE_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - French - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - French - Google" SEC_FRE_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - French - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Hindi - Google" SEC_HIN_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Hindi - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Hungarian - Deepl" SEC_HUN_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Hungarian - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - Hungarian - Google" SEC_HUN_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Hungarian - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Indonesian - Deepl" SEC_IND_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Indonesian - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - Indonesian - Google" SEC_IND_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Indonesian - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Italian - Deepl" SEC_ITA_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Italian - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - Italian - Google" SEC_ITA_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Italian - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Japanese - Deepl" SEC_JAP_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Japanese - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - Japanese - Google" SEC_JAP_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Japanese - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Nepali - Google" SEC_NEP_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Nepali - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Persian - Google" SEC_PER_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Persian - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Polish - Deepl" SEC_POL_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Polish - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - Polish - Google" SEC_POL_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Polish - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Punjabi - Google" SEC_PUN_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Punjabi - Google.lnk"
  SectionEnd
  
  
  Section /o "Machine Translation - Romanian - Deepl" SEC_RD
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Romanian - Deepl.lnk"
  SectionEnd

;  Section /o "Machine Translation - Romanian - Deepl" SEC_RD
;	  SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
;	  
;	  ; Check if the file already exists at the destination
;	  IfFileExists "$APPDATA\Microsoft\Windows\SendTo\machine-translate-docx.exe - Romanian - Deepl.lnk" checkFile
;	  
;	  ; If the file doesn't exist, install it
;	  File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Romanian - Deepl.lnk"
;	  Goto done
;
;	  checkFile:
;	  ; Get timestamps for both source and destination files
;	  GetFileTime "$APPDATA\Microsoft\Windows\SendTo\machine-translate-docx.exe - Romanian - Deepl.lnk" $0 $1
;	  GetFileTime "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Romanian - Deepl.lnk" $2 $3
;	  
;	  ; Compare the timestamps
;	  IntCmp $0 $2 newerFile skipReplace
;	  IntCmp $1 $3 newerFile skipReplace
;	  
;	  newerFile:
;	  ; Ask the user if they want to replace the file
;	  MessageBox MB_YESNO "A newer shortcut file for 'machine-translate-docx.exe - Romanian - Deepl' already exists. Do you want to replace it?" IDYES replaceFile
;	  Goto done
;	  
;	  replaceFile:
;	  File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Romanian - Deepl.lnk"
;	  
;	  skipReplace:
;	  ; File is not newer or user chose not to replace
;	  
;	  done:
;	SectionEnd
Var FontHandle

Function .onInit
    ; Create a bold Arial font with size 16 for the title
    System::Call 'gdi32::CreateFont(i 16, i 0, i 0, i 0, i 700, i 0, i 0, i 0, i 1, i 0, i 0, i 0, i 0, t "Arial") i .r0'
    StrCpy $FontHandle $0  ; Save the font handle to $FontHandle
FunctionEnd

Function InfoPage1
    nsDialogs::Create 1018
    Pop $0
    ${If} $0 == error
        Abort
    ${EndIf}

    ; Create a label for the title with size 16 and bold font
    ${NSD_CreateLabel} 10u 10u 100% 20u "Search and Replace Excel Files"
    Pop $1  ; Get the control handle for the title
    SendMessage $1 ${WM_SETFONT} $FontHandle 0  ; Apply the custom bold font (size 16) to the title

    ; Create a text block for the information
    ${NSD_CreateLabel} 10u 40u 100% 40u "The Excel search and replace files will be overwritten. Please backup all Excel files you modified in the $INSTDIR folder before installing."
    Pop $2  ; Get the control handle for the text block

    ; Create a clickable link for documentation
    ${NSD_CreateLink} 10u 90u 100% 12u "View Documentation"
    Pop $3  ; Get the control handle for the link
    SendMessage $3 ${WM_SETFONT} 0 0  ; Reset font to default for this link
	
    ; Use NSD_OnClick for more direct event handling
    ${NSD_OnClick} $3 LinkClicked

    nsDialogs::Show
FunctionEnd

Function LinkClicked
    ExecShell "open" "https://github.com/translation-robot/machine-translate-docx/"
FunctionEnd

  Section /o "Machine Translation - Romanian - Google" SEC_ROM_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Romanian - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Russian - Deepl" SEC_RUS_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Russian - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - Russian - Google" SEC_RUS_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Russian - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Spanish - Deepl" SEC_SPA_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Spanish - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - Spanish - Google" SEC_SPA_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Spanish - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Telugu - Google" SEC_TEL_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Telugu - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Thai - Google" SEC_THA_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Thai - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Ukrainian - Deepl" SEC_UKR_DEEPL
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Ukrainian - Deepl.lnk"
  SectionEnd

  Section /o "Machine Translation - Ukrainian - Google" SEC_UKR_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Ukrainian - Google.lnk"
  SectionEnd

  Section /o "Machine Translation - Urdu - Google" SEC_URD_GOOGLE
    SetOutPath "$APPDATA\Microsoft\Windows\SendTo"
    File "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Urdu - Google.lnk"
  SectionEnd

SectionGroupEnd ; Close the Shortcuts section group

!define MUI_COMPONENTS_PAGE_UNFOLDERED "Shortcuts"

!define MUI_UNFINISH_TITLE "Uninstall ${APP_NAME}"
!define MUI_UNFINISH_SUBTITLE "This will remove ${APP_NAME} from your system."

Section "Uninstall"
  RMDir /r "$INSTDIR\bin"
  RMDir /r "$INSTDIR\source_code"
  RMDir /r "$INSTDIR\ConEmuPack"
  Delete "$DESKTOP\Machine Translation Robot.lnk"
  Delete "$APPDATA\Microsoft\Windows\SendTo\Machine Translation Robot.lnk"
  Delete "$APPDATA\Microsoft\Windows\SendTo\machine-translate-docx.exe Graphical Inface - Ask language.lnk"
  Delete "$APPDATA\Microsoft\Windows\SendTo\machine-translate-docx.exe Show version.lnk"
  Delete "$APPDATA\Microsoft\Windows\SendTo\machine-translate-docx.exe Split translation - any language.lnk"
  
  
  ; Remove registry entries; Remove from Local Machine (if installed for all users)
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"
	; Remove from Current User (if installed for the current user only)
	DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"	
SectionEnd

Function .onInstSuccess
  MessageBox MB_YESNO|MB_ICONQUESTION "Would you like to launch ${APP_NAME} now?" IDYES Launch IDNO NoLaunch
  
  Launch:
    Exec "$INSTDIR\bin\machine_translate_gui.exe"
	Exec "cmd /c start http://github.com/translation-robot/machine-translate-docx/"
  NoLaunch:
  
  
  ;Exec "cmd /c start http://github.com/translation-robot/machine-translate-docx/"
  ExecShell "open" "http://github.com/translation-robot/machine-translate-docx/"
FunctionEnd

; Pages
