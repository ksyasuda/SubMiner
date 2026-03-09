!include "MUI2.nsh"
!include "nsDialogs.nsh"

Var WindowsMpvShortcutStartMenuPath
Var WindowsMpvShortcutDesktopPath

!macro ResolveWindowsMpvShortcutPaths
  !ifdef MENU_FILENAME
    StrCpy $WindowsMpvShortcutStartMenuPath "$SMPROGRAMS\${MENU_FILENAME}\SubMiner mpv.lnk"
  !else
    StrCpy $WindowsMpvShortcutStartMenuPath "$SMPROGRAMS\SubMiner mpv.lnk"
  !endif
  StrCpy $WindowsMpvShortcutDesktopPath "$DESKTOP\SubMiner mpv.lnk"
!macroend

!ifndef BUILD_UNINSTALLER
Var WindowsMpvShortcutStartMenuCheckbox
Var WindowsMpvShortcutDesktopCheckbox
Var WindowsMpvShortcutStartMenuEnabled
Var WindowsMpvShortcutDesktopEnabled
Var WindowsMpvShortcutDefaultsInitialized

!macro customInit
  StrCpy $WindowsMpvShortcutStartMenuEnabled "1"
  StrCpy $WindowsMpvShortcutDesktopEnabled "1"
  StrCpy $WindowsMpvShortcutDefaultsInitialized "0"
!macroend

!macro customPageAfterChangeDir
  PageEx custom
    PageCallbacks WindowsMpvShortcutPageCreate WindowsMpvShortcutPageLeave
    Caption " "
  PageExEnd
!macroend

Function HasExistingInstallation
  ReadRegStr $0 SHELL_CONTEXT "Software\${APP_GUID}" InstallLocation
  ${if} $0 == ""
    Push "0"
  ${else}
    Push "1"
  ${endif}
FunctionEnd

Function InitializeWindowsMpvShortcutDefaults
  ${if} $WindowsMpvShortcutDefaultsInitialized == "1"
    Return
  ${endif}

  !insertmacro ResolveWindowsMpvShortcutPaths
  Call HasExistingInstallation
  Pop $0

  ${if} $0 == "1"
    ${if} ${FileExists} "$WindowsMpvShortcutStartMenuPath"
      StrCpy $WindowsMpvShortcutStartMenuEnabled "1"
    ${else}
      StrCpy $WindowsMpvShortcutStartMenuEnabled "0"
    ${endif}

    ${if} ${FileExists} "$WindowsMpvShortcutDesktopPath"
      StrCpy $WindowsMpvShortcutDesktopEnabled "1"
    ${else}
      StrCpy $WindowsMpvShortcutDesktopEnabled "0"
    ${endif}
  ${else}
    StrCpy $WindowsMpvShortcutStartMenuEnabled "1"
    StrCpy $WindowsMpvShortcutDesktopEnabled "1"
  ${endif}

  StrCpy $WindowsMpvShortcutDefaultsInitialized "1"
FunctionEnd

Function WindowsMpvShortcutPageCreate
  Call InitializeWindowsMpvShortcutDefaults

  !insertmacro MUI_HEADER_TEXT "Windows mpv launcher" "Choose where to create the optional SubMiner mpv shortcuts."

  nsDialogs::Create 1018
  Pop $0

  ${NSD_CreateLabel} 0u 0u 300u 30u "SubMiner mpv launches SubMiner.exe --launch-mpv so people can open mpv with the SubMiner profile from a separate Windows shortcut."
  Pop $0

  ${NSD_CreateCheckbox} 0u 44u 280u 12u "Create Start Menu shortcut"
  Pop $WindowsMpvShortcutStartMenuCheckbox
  ${if} $WindowsMpvShortcutStartMenuEnabled == "1"
    ${NSD_Check} $WindowsMpvShortcutStartMenuCheckbox
  ${endif}

  ${NSD_CreateCheckbox} 0u 64u 280u 12u "Create Desktop shortcut"
  Pop $WindowsMpvShortcutDesktopCheckbox
  ${if} $WindowsMpvShortcutDesktopEnabled == "1"
    ${NSD_Check} $WindowsMpvShortcutDesktopCheckbox
  ${endif}

  ${NSD_CreateLabel} 0u 90u 300u 24u "Upgrades preserve the current SubMiner mpv shortcut locations instead of recreating shortcuts you already removed."
  Pop $0

  nsDialogs::Show
FunctionEnd

Function WindowsMpvShortcutPageLeave
  ${NSD_GetState} $WindowsMpvShortcutStartMenuCheckbox $0
  ${if} $0 == ${BST_CHECKED}
    StrCpy $WindowsMpvShortcutStartMenuEnabled "1"
  ${else}
    StrCpy $WindowsMpvShortcutStartMenuEnabled "0"
  ${endif}

  ${NSD_GetState} $WindowsMpvShortcutDesktopCheckbox $0
  ${if} $0 == ${BST_CHECKED}
    StrCpy $WindowsMpvShortcutDesktopEnabled "1"
  ${else}
    StrCpy $WindowsMpvShortcutDesktopEnabled "0"
  ${endif}
FunctionEnd

!macro customInstall
  Call InitializeWindowsMpvShortcutDefaults
  !insertmacro ResolveWindowsMpvShortcutPaths

  ${if} $WindowsMpvShortcutStartMenuEnabled == "1"
    !ifdef MENU_FILENAME
      CreateDirectory "$SMPROGRAMS\${MENU_FILENAME}"
    !endif
    CreateShortCut "$WindowsMpvShortcutStartMenuPath" "$appExe" "--launch-mpv" "$appExe" 0 "" "" "Launch mpv with the SubMiner profile"
    # electron-builder's upstream NSIS templates use the same WinShell call for AppUserModelID wiring.
    # WinShell.dll comes from electron-builder's cached nsis-resources bundle, so bun run build:win needs no extra repo-local setup.
    ClearErrors
    WinShell::SetLnkAUMI "$WindowsMpvShortcutStartMenuPath" "${APP_ID}"
  ${else}
    Delete "$WindowsMpvShortcutStartMenuPath"
  ${endif}

  ${if} $WindowsMpvShortcutDesktopEnabled == "1"
    CreateShortCut "$WindowsMpvShortcutDesktopPath" "$appExe" "--launch-mpv" "$appExe" 0 "" "" "Launch mpv with the SubMiner profile"
    # ClearErrors keeps the optional AUMI assignment non-fatal if the packaging environment is missing WinShell.
    ClearErrors
    WinShell::SetLnkAUMI "$WindowsMpvShortcutDesktopPath" "${APP_ID}"
  ${else}
    Delete "$WindowsMpvShortcutDesktopPath"
  ${endif}

  System::Call 'Shell32::SHChangeNotify(i 0x8000000, i 0, i 0, i 0)'
!macroend
!endif

!macro customUnInstall
  !insertmacro ResolveWindowsMpvShortcutPaths
  Delete "$WindowsMpvShortcutStartMenuPath"
  Delete "$WindowsMpvShortcutDesktopPath"
!macroend
