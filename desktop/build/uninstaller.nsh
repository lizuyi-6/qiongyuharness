!include "FileFunc.nsh"

; Long-path-safe removal around electron-builder's default uninstall (uninstaller.nsh:156-187 in
; app-builder-lib 26.18: customUnInstall runs before the default file removal).
;  - Upgrade (--updated): the default un.atomicRMDir renames every file into $PLUGINSDIR\old-install
;    and aborts with exit code 2 ("Failed to uninstall old application files") when a path exceeds
;    MAX_PATH; the installer then shows an MB_OK box without /SD, which blocks silent upgrades.
;    The runtime tree is re-extracted by the new installer, so only it is removed here.
;  - Real uninstall: remove the whole $INSTDIR, and the userData dir pinned by main.cjs:10
;    (%APPDATA%\QiongyuOffice, not the productName dir that deleteAppDataOnUninstall targets).
; Guards: $INSTDIR non-empty, not a drive root, and $INSTDIR\穹驭办公.exe present.
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
    StrCpy $1 "$INSTDIR\resources\qo-runtime"
  ${else}
    StrCpy $1 "$INSTDIR"
  ${endif}
  ${if} ${FileExists} "$1\*.*"
    System::Call 'kernel32::SetEnvironmentVariableW(w "QO_ELECTRON_UNINSTALL_TARGET", w "$1") i .r0'
    ${if} $0 != 0
      ExecWait '"$SYSDIR\cmd.exe" /d /c rd /s /q "\\?\%QO_ELECTRON_UNINSTALL_TARGET%"' $2
    ${endif}
  ${endif}
qoSkipLongPathRemoval:

  ${ifNot} ${isUpdated}
    RMDir /r "$APPDATA\QiongyuOffice"
  ${endif}
!macroend
