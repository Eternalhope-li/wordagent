; WordAgent AI 文档助手 — NSIS 安装脚本
Unicode true
 !define APP_NAME "WordAgent"
!define APP_VERSION "1.7.5"
!define APP_PUBLISHER "eternalhope"
!define APP_EXE "WordAgent.exe"

Name "WordAgent AI 文档助手"
OutFile "installer\WordAgent_Setup_${APP_VERSION}.exe"
InstallDir "$LOCALAPPDATA\Programs\WordAgent"
RequestExecutionLevel user
SetCompressor /SOLID lzma

!include "MUI2.nsh"
!include "FileFunc.nsh"

VIProductVersion "1.7.5.0"
 VIAddVersionKey "ProductName" "WordAgent AI 文档助手"
 VIAddVersionKey "CompanyName" "eternalhope"
 VIAddVersionKey "FileDescription" "WordAgent AI 文档助手 安装程序"
VIAddVersionKey "FileVersion" "1.7.5"
VIAddVersionKey "ProductVersion" "1.7.5"
VIAddVersionKey "LegalCopyright" "Copyright (c) 2026 eternalhope"

!define MUI_ICON "assets\icon.ico"
!define MUI_UNICON "assets\icon.ico"
!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\WordAgent.exe"
!define MUI_FINISHPAGE_RUN_TEXT "立即启动 WordAgent"
!define MUI_FINISHPAGE_LINK "访问项目主页"
!define MUI_FINISHPAGE_LINK_LOCATION "https://github.com/"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "SimpChinese"

Section "主程序" SecMain
  SetOutPath "$INSTDIR"
  File /r /x output "dist\WordAgent\*"
  ; 首次安装时生成 .env（已含 DeepSeek API 配置）
  IfFileExists "$INSTDIR\.env" env_skip 0
  FileOpen $0 "$INSTDIR\.env" w
  FileWrite $0 "DEEPSEEK_API_KEY=$\r$\n"
  FileWrite $0 "DEEPSEEK_BASE_URL=https://api.deepseek.com$\r$\n"
  FileWrite $0 "DEEPSEEK_MODEL=deepseek-v4-flash$\r$\n"
  FileWrite $0 "DEEPSEEK_TEMPERATURE=0.7$\r$\n"
  FileWrite $0 "DEEPSEEK_MAX_TOKENS=8192$\r$\n"
  FileWrite $0 "DEEPSEEK_TIMEOUT=180$\r$\n"
  FileWrite $0 "DEEPSEEK_MAX_RETRIES=3$\r$\n"
  FileWrite $0 "OUTPUT_DIR=output$\r$\n"
  FileWrite $0 "MEMORY_FILE=memory.json$\r$\n"
  FileClose $0
  env_skip:
  ; 卸载程序
  WriteUninstaller "$INSTDIR\uninstall.exe"
  ; 快捷方式
  CreateDirectory "$SMPROGRAMS\WordAgent"
  CreateShortcut "$SMPROGRAMS\WordAgent\WordAgent.lnk" "$INSTDIR\WordAgent.exe" "" "$INSTDIR\WordAgent.exe" 0
  CreateShortcut "$SMPROGRAMS\WordAgent\卸载 WordAgent.lnk" "$INSTDIR\uninstall.exe"
  CreateShortcut "$DESKTOP\WordAgent.lnk" "$INSTDIR\WordAgent.exe" "" "$INSTDIR\WordAgent.exe" 0
  ; 卸载信息（控制面板可卸载）
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\WordAgent" "DisplayName" "WordAgent AI 文档助手"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\WordAgent" "DisplayVersion" "${APP_VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\WordAgent" "Publisher" "${APP_PUBLISHER}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\WordAgent" "DisplayIcon" "$INSTDIR\WordAgent.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\WordAgent" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\WordAgent" "InstallLocation" "$INSTDIR"
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\WordAgent" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\WordAgent" "NoRepair" 1
  ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
  IntFmt $2 "0x%08X" $2
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\WordAgent" "EstimatedSize" $2
SectionEnd

Section "Uninstall"
  Delete "$DESKTOP\WordAgent.lnk"
  Delete "$SMPROGRAMS\WordAgent\WordAgent.lnk"
  Delete "$SMPROGRAMS\WordAgent\卸载 WordAgent.lnk"
  RMDir "$SMPROGRAMS\WordAgent"
  Delete "$INSTDIR\uninstall.exe"
  RMDir /r "$INSTDIR"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\WordAgent"
SectionEnd
