!include LogicLib.nsh

Var KeepData

Function .onInit
  ; Check for existing app data directory
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
  ${Else}
    StrCpy $KeepData "1"
  ${EndIf}
  initDone:
FunctionEnd

Function .onInstSuccess
  ${If} $KeepData == "0"
    RMDir /r "$APPDATA\ReportGenX"
  ${EndIf}
FunctionEnd

Function un.onInit
  ${If} ${FileExists} "$APPDATA\ReportGenX\*.*"
    MessageBox MB_YESNO|MB_ICONQUESTION \
      "检测到应用数据（报告、配置等）。$\n$\n是否保留已有数据？$\n$\n点击[是]保留，点击[否]彻底删除。" \
      IDYES keepUnData IDNO clearUnData
    clearUnData:
      StrCpy $KeepData "0"
      Goto unInitDone
    keepUnData:
      StrCpy $KeepData "1"
      Goto unInitDone
    unInitDone:
  ${Else}
    StrCpy $KeepData "1"
  ${EndIf}
FunctionEnd

Function un.onUninstSuccess
  ${If} $KeepData == "0"
    RMDir /r "$APPDATA\ReportGenX"
  ${EndIf}
FunctionEnd
