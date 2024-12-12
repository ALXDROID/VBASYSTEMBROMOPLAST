
' TechHelp Free Template Copyright (c) 2023 by Richard Rost and AccessLearningZone.com
'
' Sleep functions covered here: https://599cd.com/Sleep
'

Option Compare Database
Option Explicit

Public Function HexToRGB(H As String) As Long

    Dim R As Long, G As Long, B As Long
    
    If Left(H, 1) = "#" Then H = Right(H, Len(H) - 1)
    R = CLng("&H" & Left(H, 2))
    G = CLng("&H" & Mid(H, 3, 2))
    B = CLng("&H" & Right(H, 2))
    HexToRGB = RGB(R, G, B)
    
End Function

Public Function MyMsgBox(Prompt As String, Optional Title As String, _
        Optional Button1 As String = "OK", Optional Button2 As String = "Cancel", _
        Optional Button3 As String, Optional FormBackColor As String = "#8EA3BD", Optional Picture As String = "C:\Users\aphex\Documents\BROMOPLAST\Bromnoplast\icoAccess\log.jpg") As String

    Dim Args As String
    
    If Button1 = "" Then Button1 = "OK"
    
    Args = "Prompt=" & Prompt & ";"
    If Title <> "" Then Args = Args & "Title=" & Title & ";"
     Args = Args & "Button1=" & Button1 & ";"
      If Button2 <> "" Then Args = Args & "Button2=" & Button2 & ";"
      If Button3 <> "" Then Args = Args & "Button3=" & Button3 & ";"
     Args = Args & "FormBackColor=" & FormBackColor & ";"
     Args = Args & "Picture=" & Picture & ";"
    MsgBox Args
    DoCmd.OpenForm "MyMsgBox", WindowMode:=acDialog, OpenArgs:=Args
    
    MyMsgBox = TempVars("MyMsgBox")
    
End Function