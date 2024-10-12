[Setup]
AppId={{5418D107-9E34-4BB5-A055-C7852C20B082}
AppName=Machine Translation Robot
AppVersion=0.1
AppPublisher=Blue Sun
AppPublisherURL=https://github.com/translation-robot/machine-translate-docx
AppSupportURL=https://github.com/translation-robot/machine-translate-docx
AppUpdatesURL=https://github.com/translation-robot/machine-translate-docx
DefaultDirName=C:\SMTVRobot
DisableDirPage=yes
DefaultGroupName=Translation Robot
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputBaseFilename=mysetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Types]                                    
Name: "full"; Description: "Full installation"    
Name: "compact"; Description: "Compact installation"
Name: "custom"; Description: "Custom installation"; Flags: iscustom


[Components]
Name: "program"; Description: "Main program"; Types: custom full compact; Flags: fixed
Name: "desktopshortcut"; Description: "Desktop shortcut"; Types: custom full compact; Flags: fixed
Name: "sendto"; Description: "SendTo shortcuts for explorer"; Types: custom
Name: "sendto\show_version"; Description: "Show version"; Types: custom full compact
Name: "sendto\split_translation"; Description: "Split translation - any language"; Types: custom full compact 
Name: "sendto\bulgarian_deepl"; Description: "Bulgarian - Deepl"; Types: custom
Name: "sendto\bulgarian_google"; Description: "Bulgarian - Google"; Types: custom
Name: "sendto\traditional_chinese_google"; Description: "Chinese Traditional - Google"; Types: custom
Name: "sendto\simplified_chinese_google"; Description: "Chinese Simplified - Google"; Types: custom
Name: "sendto\chinese_simplified_deepl"; Description: "Chinese Simplified - Deepl"; Types: custom
Name: "sendto\czech_deepl"; Description: "Czech - Deepl"; Types: custom
Name: "sendto\czech_google"; Description: "Czech - Google"; Types: custom
Name: "sendto\french_deepl"; Description: "French - Deepl"; Types: custom
Name: "sendto\french_google"; Description: "French - Google"; Types: custom
Name: "sendto\hindi_google"; Description: "Hindi - Google"; Types: custom
Name: "sendto\hungarian_deepl"; Description: "Hungarian - Deepl"; Types: custom
Name: "sendto\hungarian_google"; Description: "Hungarian - Google"; Types: custom
Name: "sendto\indonesian_deepl"; Description: "Indonesian - Deepl"; Types: custom
Name: "sendto\Indonesian_google"; Description: "Indonesian - Google"; Types: custom
Name: "sendto\italian_deepl"; Description: "Italian - Deepl"; Types: custom
Name: "sendto\italian_google"; Description: "Italian - Google"; Types: custom
Name: "sendto\japanese_deepl"; Description: "Japanese - Deepl"; Types: custom
Name: "sendto\japanese_google"; Description: "Japanese - Google"; Types: custom
Name: "sendto\nepali_google"; Description: "Nepali - Google"; Types: custom
Name: "sendto\persian_google"; Description: "Persian - Google"; Types: custom
Name: "sendto\polish_deepl"; Description: "Polish - Deepl"; Types: custom
Name: "sendto\polish_google"; Description: "Polish - Google"; Types: custom
Name: "sendto\punjabi_google"; Description: "Punjabi - Google"; Types: custom
Name: "sendto\romanian_deepl"; Description: "Romanian - Deepl"; Types: custom
Name: "sendto\romanian_google"; Description: "Romanian - Google"; Types: custom
Name: "sendto\russian_deepl"; Description: "Russian - Deepl"; Types: custom
Name: "sendto\russian_google"; Description: "Russian - Google"; Types: custom
Name: "sendto\spanish_deepl"; Description: "Spanish - Deepl"; Types: custom
Name: "sendto\spanish_google"; Description: "Spanish - Google"; Types: custom
Name: "sendto\telugu_google"; Description: "Telugu - Google"; Types: custom
Name: "sendto\thai_google"; Description: "Thai - Google"; Types: custom
Name: "sendto\ukrainian_deepl"; Description: "Ukrainian - Deepl"; Types: custom
Name: "sendto\ukrainian_google"; Description: "Ukrainian - Google"; Types: custom
Name: "sendto\urdu_google"; Description: "Urdu - Google"; Types: custom


[Files]
; Mandatory program files
Source: "C:\SMTVRobot-installer\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: program
                               
     
Source: "C:\SMTVRobot-installer\source_code\SendTo\Machine Translation Robot.lnk"; DestDir: "{userdesktop}"; Components:desktopshortcut


;Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe   Graphical Inface - Ask language.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\ask_language  
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe   Show version.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\show_version
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe   Split translation - any language.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\split_translation


Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Bulgarian - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\bulgarian_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Bulgarian - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\bulgarian_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Chinese Traditional - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\traditional_chinese_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Chinese Simplidied - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\simplified_chinese_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Chinese Simplified - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\chinese_simplified_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Czech - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\czech_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Czech - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\czech_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - French - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\french_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - French - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\french_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Hindi - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\hindi_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Hungarian - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\hungarian_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Hungarian - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\hungarian_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Indonesian - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\indonesian_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Indonesian - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\Indonesian_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Italian - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\italian_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Italian - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\italian_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Japanese - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\japanese_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Japanese - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\japanese_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Nepali - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\nepali_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Persian - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\persian_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Polish - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\polish_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Polish - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\polish_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Punjabi - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\punjabi_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Romanian - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\romanian_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Romanian - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\romanian_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Russian - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\russian_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Russian - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\russian_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Spanish - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\spanish_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Spanish - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\spanish_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Telugu - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\telugu_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Thai - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\thai_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Ukrainian - Deepl.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\ukrainian_deepl
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Ukrainian - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\ukrainian_google
Source: "C:\SMTVRobot-installer\source_code\SendTo\machine-translate-docx.exe - Urdu - Google.lnk"; DestDir: "{userappdata}\Microsoft\Windows\SendTo"; Components:sendto\urdu_google


[Code]
var
  LaunchCheckboxPage: TWizardPage;
  LaunchCheckbox: TCheckBox;
  ResultCode: Integer;

procedure InitializeWizard();
begin
  // Create a custom page with a checkbox at the end of the installation process
  LaunchCheckboxPage := CreateCustomPage(wpFinished, 'Post-Installation', 'Would you like to launch the program now?');
  
  // Create a checkbox on the custom page
  LaunchCheckbox := TCheckBox.Create(LaunchCheckboxPage);
  LaunchCheckbox.Parent := LaunchCheckboxPage.Surface;
  LaunchCheckbox.Caption := 'Launch Machine Translation Robot';
  LaunchCheckbox.Left := ScaleX(10);
  LaunchCheckbox.Top := ScaleY(10);
  LaunchCheckbox.Width := ScaleX(200);
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = wpFinished then
  begin
    // Launch the program if the checkbox is checked
    if LaunchCheckbox.Checked then
    begin
      Exec(ExpandConstant('{app}\img\machine_translate_gui.exe'), '', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
    end;
  end;
end;

