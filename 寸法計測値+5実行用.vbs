'----------------------------------------------------------------------------------------------
'環境:Win7(64bit)+LT2015(dxf2000)+acadremon.dll
'ファイル2つ必要です。
'ソースファイル名実行用.vbs (このファイル:ソースファイル名の後ろに「実行用」3文字を追加する。)
'ソースファイル名.vbs     (ソース用ファイル)
'----------------------------------------------------------------------------------------------
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
srcFolder=objFSO.GetParentFolderName(WScript.ScriptFullName) & "\"
srcFile=mid(Wscript.ScriptName,1,len(Wscript.ScriptName)-3-1-3) & ".vbs"
CreateObject("WScript.Shell").Run "C:\Windows\SysWOW64\WScript.exe " & srcFolder & srcFile,0
