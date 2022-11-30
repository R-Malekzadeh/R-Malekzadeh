Attribute VB_Name = "Module1"

Public Function CheckDate(StrDate As String) As Boolean
    
    On Error GoTo myerr
    
    CheckDate = True
    
    If Len(Trim(StrDate)) < 10 Then
        CheckDate = False
        GoTo myexit
    End If
    
    
    If Val(Mid(Trim(StrDate), 3, 2)) < 60 Then
        CheckDate = False
        GoTo myexit
    End If
    
    If Val(Mid(Trim(StrDate), 6, 2)) > 12 Or Val(Mid(Trim(StrDate), 6, 2)) = 0 Then
        CheckDate = False
        GoTo myexit
    End If
    
    If Val(Mid(Trim(StrDate), 9, 2)) > 31 Or Val(Mid(Trim(StrDate), 9, 2)) = 0 Then
        CheckDate = False
        GoTo myexit
    End If
    
    If Val(Mid(Trim(StrDate), 6, 2)) > 6 And Val(Mid(Trim(StrDate), 9, 2)) > 30 Then
        CheckDate = False
        GoTo myexit
    End If
    
    GoTo myexit
    
myerr:
    CheckDate = False
    GoTo myexit
    
myexit:

End Function

'Public Function Color()
   'cdbColor.CancelError = True
   
'   On Error GoTo dbErrHandler
   
'  cdbColor.Flags = cdlCCFullOpen + cdlCCHelpButton
'ColorDB
'  cdbColor.ShowColor
'  frmMain.BackColor = cdbColor.Color
'  Exit Function
  
'dbErrHandler:
  
'  Exit Function
'End Function


