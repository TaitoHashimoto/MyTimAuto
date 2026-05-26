Dim shell, fso, script_dir
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
script_dir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.Run "pythonw """ & script_dir & "\punch.py"" 1", 0, False
