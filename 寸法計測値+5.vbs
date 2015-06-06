'----------------------------------------------------------------------------------------------
'寸法計測値+5.vbs
'環境:Win7+LT2015(dxf2000)+acadremon.dll
'ファイル2つ必要です。
'ソースファイル名実行用.vbs (ソースファイル名の後ろに「実行用」3文字を追加する。)
'ソースファイル名.vbs       (ソース用ファイル)
'Ver0.1 2015/06/06
'----------------------------------------------------------------------------------------------
Dim Acad
Call Main 
Sub Main()
    Set Acad = CreateObject("AcadRemocon.Body")
''''If Not Acad.acDxfOut("", "DWG", False) Then Er: Exit Sub 　　　　　　　　　　　　　　　　　　　　'修正箇所 
    If Not      adDxfOut("", "DWG", False) Then Er: Exit Sub
    If Not Acad.DxfExtract(Cnt,ExtArr,"ENTITIES","","DIMENSION","1|11|21|42|70") Then Er: Exit Sub
    if Cnt=0 then
       msgbox "寸法なし。OKで終了。"
       WScript.Quit
    else
    msgbox"開始###" & Cnt & "###"
      for i=1 to Cnt
if i mod 1000 =0 then 
    msgbox"通過###" & i & "###"
end if
            if ExtArr(1,i)="" Then
               atai = round(ExtArr(4,i)+5)
               ExtArr(5,i)=ExtArr(5,i) & vbCrLf & "  1" & vbCrLf & atai  
Acad.Wait 10
               call adLine(0,0,ExtArr(2,i),ExtArr(3,i))
            else
            end if
       next
    end if


    If Not Acad.DxfUpdate(ExtArr) Then Er: Exit Sub
''''If Not Acad.acDxfIn() Then Er: Exit Sub 　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'修正箇所 
    If Not      adDxfIn() Then Er: Exit Sub
    If Not Acad.acPostCommand("ERASE P^M^M") Then Er: Exit Sub
    
    msgbox"終了"
End Sub
'-----------------------------------------------------------------------------------------------
sub adLine(X1,Y1,X2,Y2):Acad.acPostCommand "LINE " & x1 & "," & y1 & " " & x2 & "," & y2 & "^M^M":End Sub
'-----------------------------------------------------------------------------------------------
Sub Er()
    If Acad.ErrNumber = vbObjectError + 1000 Then
    Else
       Acad.ShowError
    End If
End Sub
'========== ***** ↓↓↓↓↓ ***** =============================================================
dim CommonDxfFullPath  'AcadRemocon.dxfに固定
dim CommonDxfOpenCount 'AcadRemocon.dxfを開く回数
'-----------------------------------------------------------------------------------------------
Function adDxfOut(Message,SelectString,CanSelectLockedLayer) 'CommonDxfFullPath使用
   CommonDxfOpenCount=CommonDxfOpenCount+1
   if CommonDxfOpenCount=1 then
      call CommonDxfFullPathDef
   end if
'
    if adFileAri(CommonDxfFullPath)=1 then
       call adDeleteFile(CommonDxfFullPath)
    end if
    icheck=0
    do 
       if adFileAri(CommonDxfFullPath)=0 then exit do
       icheck=icheck+1
       if icheck=60 then
          msgbox "再実行して下さい。OKで終了。"
          WScript.Quit
       end if
       WScript.Sleep 1*1000
    loop
'
    icheck=0
    if SelectString="DWG" then
           Acad.acPostCommand("^C^CSELECT ALL  ")
           Acad.acPostCommand("DWGTITLED ")
           Acad.acPostCommand("FILEDIA 0 ")
           Acad.acPostCommand("DXFOUT^M")
           Acad.acPostCommand(CommonDxfFullPath&"^M")
           Acad.acPostCommand("V^M")
           Acad.acPostCommand("2000^M")
           Acad.acPostCommand("16^M")
           Acad.acPostCommand("Y^M")
           Acad.acPostCommand("FILEDIA 1 ")
           adDxfOut=true
         do
           icheck=icheck+1
           if adFileAri(CommonDxfFullPath)=1 then exit do
           if icheck=6000 then
              msgbox "再実行して下さい。OKで終了。"
              WScript.Quit
           end if
           WScript.Sleep 1*1000
        loop
'msgbox "okasii-chk###" & icheck
    elseif SelectString="" then
           Acad.acSendCommand "^C^CSELECT ","オブジェクトを選択"
           call adDialog '次の行に走るのを一時的に停止させる。
           Acad.acPostCommand("DWGTITLED ")
           Acad.acPostCommand("FILEDIA 0 ")
           Acad.acPostCommand("DXFOUT^M")
           Acad.acPostCommand(CommonDxfFullPath&"^M")
           Acad.acPostCommand("o p  ")
           Acad.acPostCommand("V^M")
           Acad.acPostCommand("2000^M")
           Acad.acPostCommand("16^M")
           Acad.acPostCommand("Y^M")
           Acad.acPostCommand("FILEDIA 1 ")
           adDxfOut=true
        do
           icheck=icheck+1
           if adFileAri(CommonDxfFullPath)=1 then exit do
           if icheck=60 then
              msgbox "再実行して下さい。adDxfOut-OKで終了。"
              WScript.Quit
           end if
         loop
    else
           MsgBox "adDxfOut-未対応" 
           WScript.Quit
    end if
end Function
'-----------------------------------------------------------------------------------------------
sub CommonDxfFullPathDef 'CommonDxfFullPath使用
    Dim objNetWork
    Set objNetWork = WScript.CreateObject("WScript.Network")
    CommonDxfFullPath="C:\Users\" & objNetWork.UserName & "\AppData\Local\Temp\AcadRemocon.dxf"
    Set objNetWork = Nothing
end sub
'-----------------------------------------------------------------------------------------------
Function adDxfIn() 'CommonDxfFullPath使用
    Acad.acPostCommand("_-INSERT ")
    Acad.acPostCommand("*" & CommonDxfFullPath & "^M")
    Acad.acPostCommand("0,0^M")
    Acad.acPostCommand("1^M")
    Acad.acPostCommand("0^M")
    adDxfIn=true
end Function
'-----------------------------------------------------------------------------------------------
Sub adDialog()
    Acad.dlLoad "オブジェクトを選択後,OKして下さい。"
    Acad.dlAddLabel "","",15
    Acad.dlShow
    Do
        Acad.dlWaitEvent CtrlName
        Select Case CtrlName
          Case "cmdOK"    : Acad.dlUnload:exit sub
          Case "cmdCancel": Acad.dlUnload:exit do
        End Select
    Loop While True
    Acad.acPostCommand("^C^C")
    WScript.Quit
end sub
'-----------------------------------------------------------------------------------------------
Function adFileAri(FullPath)
    Dim Fso
    Set Fso = CreateObject("Scripting.FileSystemObject")
    If Not Fso.FileExists(FullPath) Then
       adFileAri=0
    else
       adFileAri=1
    end if
    Set Fso = Nothing
end Function
'-----------------------------------------------------------------------------------------------
sub adDeleteFile(FullPath)
    Dim objFileSys
    Dim strDeleteFrom
    Set objFileSys = CreateObject("Scripting.FileSystemObject")
    strDeleteFrom = objFileSys.BuildPath(strScriptPath,FullPath)
    objFileSys.DeleteFile strDeleteFrom, True
    Set objFileSys = Nothing
end sub
'========== ***** ↑↑↑↑↑ ***** =============================================================
