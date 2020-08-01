Attribute VB_Name = "ModFunction"
Option Explicit
Function LireFichierTexteParLigne(Monfichier As String)
    '
    Application.ScreenUpdating = False
    On Error GoTo CodeErreur
    
    Dim IndexFichier As Integer
    'Dim Monfichier As String
    Dim ContenuLigne As String
    Dim sqlString As String
    
    IndexFichier = FreeFile()
    
    Open Monfichier For Input As #IndexFichier ' Open the file
    
    While Not EOF(IndexFichier) '
        Line Input #IndexFichier, ContenuLigne
        sqlString = sqlString & ContenuLigne
        
    Wend
    
    Close #IndexFichier ' ferme le fichier
    
    LireFichierTexteParLigne = sqlString
    Application.ScreenUpdating = True
    Exit Function
    
CodeErreur:
    MsgBox "Une erreur s'est produite..."
    Application.ScreenUpdating = True
End Function
 Function getRole(listRole As String) As String
    Dim Listdata As Variant
    Dim i As Integer
    
    Listdata = Split(listRole, ",")
    
    getRole = ""
    For i = LBound(Listdata, 1) To UBound(Listdata, 1)
        If Len(Listdata(i)) > 0 Then
            getRole = getRole & Chr(32) & "(CustRole:" & """" & Listdata(i) & """" & Chr(32) & ")"
        Else
            getRole = ""
        End If
    Next i
    
    Debug.Print getRole
 End Function

'Function getFielValue(fieldname As String, fld As String) As String
'     getFielValue = IIf(Len(Trim(rst.Fields(fld))) > 0, "(CustId:" & """" & Trim(rst.Fields("CUSTID")) & """", Chr(0))
'End Function
Function getRating(agcy As String, rating As String, risktype As String, Role As String) As String
    
    If Len(Trim(rating)) > 0 Then
        getRating = "(Type:" & """" & UCase(risktype) & """" & Chr(32) & "Name:" _
                   & """" & UCase(agcy) & """" & Chr(32) & "CustRole:" & """" & UCase(Role) & """" & Chr(32) & "Value:" & """" & UCase(rating) & """" & Chr(32) & ")"
    Else
        '   We put NR if not rated, for statistical purpose
        getRating = "NR"
    End If
        

End Function
Function getData(sfield As String, stext As String, rst As Recordset) As Variant
   getData = IIf(Len(Trim(rst.Fields(sfield))) > 0, stext & ":" & """" & Trim(rst.Fields(sfield)) & """", "")
End Function
