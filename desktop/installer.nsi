Unicode true
Name "穹驭办公"
OutFile "dist/QiongyuOffice-Setup-x64.exe"
InstallDir "$LOCALAPPDATA\Programs\QiongyuOffice"
RequestExecutionLevel user
SetCompressor /SOLID lzma
Icon "office.ico"
UninstallIcon "office.ico"
VIProductVersion "0.1.0.0"
VIAddVersionKey "ProductName" "Qiongyu Office"
VIAddVersionKey "FileDescription" "Qiongyu Office Windows Installer"
VIAddVersionKey "FileVersion" "0.1.0"
VIAddVersionKey "CompanyName" "Qiongyu Technology"

!include "MUI2.nsh"
!include "x64.nsh"
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "SimpChinese"

Function .onInit
  ${IfNot} ${RunningX64}
    MessageBox MB_OK "请在 64 位 Windows 上安装穹驭办公。"
    Abort
  ${EndIf}
FunctionEnd

Section "穹驭办公"
  SetShellVarContext current
  SetOutPath "$INSTDIR"
  File /r "dist/win-unpacked/*"
  File /oname=office.ico "office.ico"
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  CreateShortcut "$SMPROGRAMS\穹驭办公.lnk" "$INSTDIR\穹驭办公.exe" "" "$INSTDIR\office.ico"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\QiongyuOffice" "DisplayName" "穹驭办公"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\QiongyuOffice" "DisplayVersion" "0.1.0"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\QiongyuOffice" "Publisher" "Qiongyu Technology"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\QiongyuOffice" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\QiongyuOffice" "DisplayIcon" "$INSTDIR\office.ico"
SectionEnd

Section "Uninstall"
  SetShellVarContext current
  Delete "$SMPROGRAMS\穹驭办公.lnk"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\QiongyuOffice"
  RMDir /r "$LOCALAPPDATA\Programs\QiongyuOffice"
  RMDir /r "$APPDATA\QiongyuOffice"
SectionEnd
