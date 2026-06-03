!include LogicLib.nsh
!include nsDialogs.nsh
!pragma warning disable 6001

Var ChkHandle

!macro customPageAfterChangeDir
  Page custom createDataOptionsPage leaveDataOptionsPage "应用数据"

  Function createDataOptionsPage
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
    Pop $ChkHandle
    ${NSD_Uncheck} $ChkHandle
    nsDialogs::Show
  FunctionEnd

  Function leaveDataOptionsPage
    ${NSD_GetState} $ChkHandle $0
    ${If} $0 == ${BST_CHECKED}
      RMDir /r "$APPDATA\ReportGenX"
    ${EndIf}
  FunctionEnd
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
