!include "FileFunc.nsh"

!ifndef BUILD_UNINSTALLER
  !define MUI_FINISHPAGE_TITLE "穹驭办公安装完成"
  !define MUI_FINISHPAGE_TEXT "穹驭办公已安装完成，可以从桌面或开始菜单的快捷方式启动。"
  !define MUI_FINISHPAGE_RUN_TEXT "立即启动穹驭办公"
!endif

!macro customInstallMode
  StrCpy $isForceCurrentInstall "1"
!macroend

!macro customWelcomePage
  !define MUI_WELCOMEPAGE_TITLE "欢迎安装穹驭办公"
  !define MUI_WELCOMEPAGE_TEXT "穹驭办公帮助企业员工完成日常办公任务。$\r$\n安装到当前用户目录，无需管理员权限，约 1 分钟。安装前请关闭正在运行的穹驭办公。"
  !insertmacro MUI_PAGE_WELCOME
!macroend

!macro customInstall
  ${ifNot} ${FileExists} "$INSTDIR\resources\qo-runtime.7z"
    Goto qoRuntimeInstallFailed
  ${endif}

  ${ifNot} ${Silent}
    SetDetailsPrint textonly
  ${endif}
  DetailPrint "正在安装办公运行环境…"

  SetOutPath $TEMP
  ${if} ${FileExists} "$INSTDIR\resources\qo-runtime\*.*"
    System::Call 'kernel32::SetEnvironmentVariableW(w "QO_ELECTRON_UNINSTALL_TARGET", w "$INSTDIR\resources\qo-runtime") i .r0'
    ${if} $0 == 0
      Goto qoRuntimeInstallFailed
    ${endif}
    nsExec::Exec '"$SYSDIR\cmd.exe" /d /c rd /s /q "\\?\%QO_ELECTRON_UNINSTALL_TARGET%"'
    Pop $2
    ${if} $2 != "0"
      Goto qoRuntimeInstallFailed
    ${endif}
  ${endif}

  CreateDirectory "$INSTDIR\resources\qo-runtime"
  SetOutPath "$INSTDIR\resources\qo-runtime"
  Nsis7z::ExtractWithDetails "$INSTDIR\resources\qo-runtime.7z" "正在安装办公运行环境 %s"
  SetOutPath "$INSTDIR"
  Delete "$INSTDIR\resources\qo-runtime.7z"

  ${ifNot} ${FileExists} "$INSTDIR\resources\qo-runtime\node\node.exe"
    Goto qoRuntimeInstallFailed
  ${endif}
  ${ifNot} ${FileExists} "$INSTDIR\resources\qo-runtime\harness\package.json"
    Goto qoRuntimeInstallFailed
  ${endif}
  ${ifNot} ${FileExists} "$INSTDIR\resources\qo-runtime\harness\office.patch.yml"
    Goto qoRuntimeInstallFailed
  ${endif}
  ${ifNot} ${FileExists} "$INSTDIR\resources\qo-runtime\harness\office-image.mjs"
    Goto qoRuntimeInstallFailed
  ${endif}
  ${ifNot} ${FileExists} "$INSTDIR\resources\qo-runtime\harness\office-context.mjs"
    Goto qoRuntimeInstallFailed
  ${endif}

  ${ifNot} ${Silent}
    SetDetailsPrint none
  ${endif}

  !ifdef APP_INSTALLER_STORE_FILE
    ${if} $installMode == "all"
      SetShellVarContext current
    ${endif}
    Delete "$LOCALAPPDATA\${APP_INSTALLER_STORE_FILE}"
    RMDir "$LOCALAPPDATA\qiongyu-office-desktop-updater"
    ${if} $installMode == "all"
      SetShellVarContext all
    ${endif}
  !endif
  Goto qoRuntimeInstallComplete

qoRuntimeInstallFailed:
  MessageBox MB_OK|MB_ICONSTOP "运行环境解压失败，请确认磁盘空间后重新运行安装程序" /SD IDOK
  SetErrorLevel 2
  Quit
qoRuntimeInstallComplete:
!macroend

; 安装升级时仅删除旧运行时；完整卸载由 electron-builder 随后清理其余安装文件。
; 长路径通过环境变量传给 cmd，避免在命令行直接展开 Unicode 安装路径。
; 守卫：安装目录非空、不是盘符根目录、且应用 exe 存在。
!macro customUnInstall
  StrCmp "$INSTDIR" "" qoSkipLongPathRemoval
  ${GetRoot} "$INSTDIR" $0
  StrCmp "$INSTDIR" "$0" qoSkipLongPathRemoval
  StrCmp "$INSTDIR\" "$0" qoSkipLongPathRemoval
  ${ifNot} ${FileExists} "$INSTDIR\穹驭办公.exe"
    Goto qoSkipLongPathRemoval
  ${endif}

  SetOutPath $TEMP
  ${if} ${isUpdated}
    Delete "$INSTDIR\resources\qo-runtime.7z"
    StrCpy $1 "$INSTDIR\resources\qo-runtime"
  ${else}
    StrCpy $1 "$INSTDIR"
  ${endif}
  ${if} ${FileExists} "$1\*.*"
    System::Call 'kernel32::SetEnvironmentVariableW(w "QO_ELECTRON_UNINSTALL_TARGET", w "$1") i .r0'
    ${if} $0 != 0
      nsExec::Exec '"$SYSDIR\cmd.exe" /d /c rd /s /q "\\?\%QO_ELECTRON_UNINSTALL_TARGET%"'
      Pop $2
    ${endif}
  ${endif}
qoSkipLongPathRemoval:

  ${ifNot} ${isUpdated}
    RMDir /r "$APPDATA\QiongyuOffice"
    !ifdef APP_INSTALLER_STORE_FILE
      Delete "$LOCALAPPDATA\${APP_INSTALLER_STORE_FILE}"
    !endif
    RMDir /r "$LOCALAPPDATA\qiongyu-office-desktop-updater"
  ${endif}
!macroend
