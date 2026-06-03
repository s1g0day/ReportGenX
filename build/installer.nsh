!include LogicLib.nsh

!macro customInit
  ${If} ${FileExists} "$APPDATA\ReportGenX\config.yaml"
    MessageBox MB_YESNO|MB_ICONQUESTION \
      "检测到已有应用数据（报告、配置等）。$\n$\n是否保留已有数据？$\n$\n点击[是]保留，点击[否]清空后全新安装。" \
      IDYES keepData IDNO clearData
    clearData:
      RMDir /r "$APPDATA\ReportGenX"
    keepData:
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
