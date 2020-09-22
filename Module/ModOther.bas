Attribute VB_Name = "ModOther"
Private Declare Function SHPathFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Function FileExists(Path$) As Boolean
   '  MsgBox Path$
   On Error Resume Next
    If Len(Trim(Path$)) = 0 Then FileExists = False: Exit Function
       FileExists = Dir(Trim(Path$), vbNormal) <> ""
       On Error GoTo 0
End Function

Public Function PathExists(sPath As String) As Boolean
    If Len(Environ$("OS")) Then
        PathExists = CBool(SHPathFileExists(StrConv(LTrim$(sPath), vbUnicode)))
    Else
        PathExists = CBool(SHPathFileExists(LTrim$(sPath)))
    End If
End Function

Public Sub SplitPath(FullPath As String, _
                     Optional Drive As String, _
                     Optional Path As String, _
                     Optional Filename As String, _
                     Optional File As String, _
                     Optional Extension As String)
                     
 Dim nPos As Integer
 nPos = InStrRev(FullPath, "\")
 If nPos > 0 Then
   If Left$(FullPath, 2) = "\\" Then
    If nPos = 2 Then
     Drive = FullPath: Path = vbNullString: Filename = vbNullString: File = vbNullString
     Extension = vbNullString
     Exit Sub
    End If
   End If
   Path = Left$(FullPath, nPos - 1)
   Filename = Mid$(FullPath, nPos + 1)
   nPos = InStrRev(Filename, ".")
   If nPos > 0 Then
     File = Left$(Filename, nPos - 1)
     Extension = Mid$(Filename, nPos + 1)
    Else
     File = Filename
     Extension = vbNullString
   End If
  Else
   nPos = InStrRev(FullPath, ":")
   If nPos > 0 Then
     Path = Mid(FullPath, 1, nPos - 1): Filename = Mid(FullPath, nPos + 1)
     nPos = InStrRev(Filename, ".")
     If nPos > 0 Then
       File = Left$(Filename, nPos - 1)
       Extension = Mid$(Filename, nPos + 1)
      Else
       File = Filename
       Extension = vbNullString
     End If
    Else
     Path = vbNullString: Filename = FullPath
     nPos = InStrRev(Filename, ".")
     If nPos > 0 Then
       File = Left$(Filename, nPos - 1)
       Extension = Mid$(Filename, nPos + 1)
      Else
       File = Filename
       Extension = vbNullString
     End If
   End If
 End If
 If Left$(Path, 2) = "\\" Then
   nPos = InStr(3, Path, "\")
   If nPos Then
     Drive = Left$(Path, nPos - 1)
    Else
     Drive = Path
   End If
  Else
   If Len(Path) = 2 Then
    If Right$(Path, 1) = ":" Then
     Path = Path & "\"
    End If
   End If
   If Mid$(Path, 2, 2) = ":\" Then
    Drive = Left$(Path, 2)
   End If
 End If
End Sub

'form on top
Public Sub FormOnTop(f_form As Form, I As Boolean)
   
   If I = True Then 'On
      SetWindowPos f_form.hWnd, -1, 0, 0, 0, 0, &H2 + &H1
   Else      'off
      SetWindowPos f_form.hWnd, -2, 0, 0, 0, 0, &H2 + &H1
   End If
   
End Sub

Public Sub SplitRGB(ByVal lColor As Long, ByRef lRed As Long, ByRef lGreen As Long, ByRef lBlue As Long)
   lRed = lColor And &HFF
   lGreen = (lColor And &HFF00&) \ &H100&
   lBlue = (lColor And &HFF0000) \ &H10000
End Sub

