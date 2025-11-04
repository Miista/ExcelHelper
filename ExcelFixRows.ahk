#Requires AutoHotkey v2.0

#HotIf WinActive("Create Work Item.xlsx ahk_exe EXCEL.EXE ahk_class XLMAIN")
F4::
{
    Send "^a"
    Send "!hw"
    Send "{Up}{Down}"
}