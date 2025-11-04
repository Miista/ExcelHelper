$ahk2exe = "C:\Users\75224\AppData\Local\Programs\AutoHotkey\Compiler\Ahk2Exe.exe"
$baseFile = "C:\Users\75224\AppData\Local\Programs\AutoHotkey\v2\AutoHotkey32.exe"
Start-Process "$ahk2exe" @("/in ExcelHelper.ahk", "/icon _bffee642-d003-48d5-be51-09f099e91082.ico", "/compress 1", "/base $baseFile")