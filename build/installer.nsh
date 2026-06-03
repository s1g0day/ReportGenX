!include LogicLib.nsh

Var KeepData

!macro customInit
  ${If} ${FileExists} "$APPDATA\ReportGenX\*.*"
    MessageBox MB_YESNO|MB_ICONQUESTION \
      "检测到已有应用数据（报告、配置等）。$\n$\n是否保留已有数据？$\n$\n点击[是]保留，点击[否]清空后全新安装。" \
      IDYES keepData IDNO clearData
    clearData:
      StrCpy $KeepData "0"
      Goto initDone
    keepData:
      StrCpy $KeepData "1"
      Goto initDone
    initDone:
  ${Else}
    StrCpy $KeepData "1"
  ${EndIf}
!macroend

!macro customInstall
  ${If} $KeepData == "0"
    RMDir /r "$APPDATA\ReportGenX"
  ${EndIf}
!macroend

!macro customUnInit
  ${If} ${FileExists} "$APPDATA\ReportGenX\*.*"
    MessageBox MB_YESNO|MB_ICONQUESTION \
      "检测到应用数据（报告、配置等）。$\n$\n是否保留已有数据？$\n$\n点击[是]保留，点击[否]彻底删除。" \
      IDYES keepUnData IDNO clearUnData
    clearUnData:
      RMDir /r "$APPDATA\ReportGenX"
    keepUnData:
  ${EndIf}
!macroend
