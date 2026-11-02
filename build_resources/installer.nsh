; 仅向 .docx 右键菜单添加“用 Word 公式修复工具打开”，
; 不写入 .docx 默认关联，避免抢占 Microsoft Word。
!macro customInstall
  WriteRegStr SHCTX "Software\Classes\SystemFileAssociations\.docx\shell\WordMathRepair" "" "用 Word 公式修复工具打开"
  WriteRegStr SHCTX "Software\Classes\SystemFileAssociations\.docx\shell\WordMathRepair" "Icon" "$INSTDIR\${APP_EXECUTABLE_FILENAME},0"
  WriteRegStr SHCTX "Software\Classes\SystemFileAssociations\.docx\shell\WordMathRepair\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
!macroend

!macro customUnInstall
  DeleteRegKey SHCTX "Software\Classes\SystemFileAssociations\.docx\shell\WordMathRepair"
!macroend
