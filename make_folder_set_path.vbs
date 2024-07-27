Dim fs
Set fs = WScript.CreateObject("Scripting.FileSystemObject")
fs.CreateFolder("C:\\Users\\user\\path")
fs.CreateFolder("C:\\Users\\user\\git")
fs.CreateFolder("C:\\Users\\user\\git\\document")
fs.CreateFolder("C:\\Users\\user\\git\\excel_vba")
fs.CreateFolder("C:\\Users\\user\\git\\py")
fs.CreateFolder("C:\\Users\\user\\git\\vbs")
fs.CreateFolder("C:\\Users\\user\\git\\windows_macro")
fs.CreateFolder("C:\\Users\\user\\anaconda3\\Scripts")

' 管理者として実行すること
Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("SYSTEM")
strNewPath = "C:\Users\\user\path"
strCurrentPath = objEnv("PATH")
' パスがすでに存在するかチェック
If InStr(1, strCurrentPath, strNewPath, vbTextCompare) = 0 Then
    ' 新しいパスを追加
    objEnv("PATH") = strCurrentPath & ";" & strNewPath
    WScript.Echo "新しいパスが追加されました: " & strNewPath
Else
    WScript.Echo "パスは既に存在します: " & strNewPath
End If
strNewPath = "C:\Users\\user\anaconda3\Scripts"
strCurrentPath = objEnv("PATH")
' パスがすでに存在するかチェック
If InStr(1, strCurrentPath, strNewPath, vbTextCompare) = 0 Then
    ' 新しいパスを追加
    objEnv("PATH") = strCurrentPath & ";" & strNewPath
    WScript.Echo "新しいパスが追加されました: " & strNewPath
Else
    WScript.Echo "パスは既に存在します: " & strNewPath
End If