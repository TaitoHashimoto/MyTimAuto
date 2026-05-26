Dim shell, fso, script_dir
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
script_dir = fso.GetParentFolderName(WScript.ScriptFullName)
' pythonw.exe を PATH から呼び出すことで、ユーザー名やバージョンに依存しない
shell.Run "pythonw """ & script_dir & "\overtime.py""", 0, False
