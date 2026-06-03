!include LogicLib.nsh
!include nsDialogs.nsh
!pragma warning disable 6001

Var KeepData
Var KeepDataCheckbox

!macro customInit
  ; Init data preservation state (default: keep)
  StrCpy $KeepData "1"
!macroend

!macro customPageAfterChangeDir
  Page custom createDataOptionsPage leaveDataOptionsPage "应用数据"

  Function createDataOptionsPage
    ; Only show this page if existing app data is found
    ${IfNot} ${FileExists} "$APPDATA\ReportGenX\config.yaml"
      Abort
    ${EndIf}

    nsDialogs::Create 1018
    Pop $0
    ${If} $0 == error
      Abort
    ${EndIf}

    ${NSD_CreateLabel} 0 0 100% 12u "检测到已有应用数据（报告、配置等）。"
    Pop $0

    ${NSD_CreateCheckbox} 0 20u 100% 12u "删除已有应用数据（含报告、配置、数据库）"
    Pop $KeepDataCheckbox
    ${NSD_Uncheck} $KeepDataCheckbox

    nsDialogs::Show
  FunctionEnd

  Function leaveDataOptionsPage
    ${NSD_GetState} $KeepDataCheckbox $0
    ${If} $0 == ${BST_CHECKED}
      StrCpy $KeepData "0"
    ${Else}
      StrCpy $KeepData "1"
    ${EndIf}
  FunctionEnd
!macroend

!macro customInstall
  ${If} $KeepData == "0"
    RMDir /r "$APPDATA\ReportGenX"
  ${EndIf}
!macroend

!macro customUnInit
  ${IfNot} ${Silent}
    ${If} ${FileExists} "$APPDATA\ReportGenX\config.yaml"
      MessageBox MB_YESNO|MB_ICONQUESTION \
        "检测到应用数据（报告、配置等）。$\n$\n是否保留已有数据？$\n$\n点击[是]保留，点击[否]彻底删除。" \
        IDYES keepUnData IDNO clearUnData
      clearUnData:
        RMDir /r "$APPDATA\ReportGenX"
      keepUnData:
    ${EndIf}
  ${EndIf}
!macroend
