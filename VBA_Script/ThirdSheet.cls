VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Third()

Dim SessionUsing As String
Dim SessionNotUsing As String
Dim Locking As String
Dim Unlocking As String


lastRow = Sheet1.Cells(Rows.Count, 7).End(xlUp).Row
    ' Call Sheet2 Result Information
    For I = 1 To lastRow
        Cells(I, 2) = Sheet2.Cells(I, 7)
        Cells(I, 3) = Sheet2.Cells(I, 8)
        Cells(I, 4) = Sheet2.Cells(I, 9)
        Cells(I, 5) = Sheet2.Cells(I, 10)
    Next

    
    SessionUsing = "########################"
    Unlocking = "#######################"
    
    Locking = "#######################"
    SessionNotUsing = "########################"
  
    
    For I = 1 To lastRow
    
        ' Can Use
        If StrComp(Cells(I, 5), SessionNotUsing) = 0 Or StrComp(Cells(I, 5), Locking) = 0 Then
            ' Green Color
            Cells(I, 4).Interior.ColorIndex = 4
        
        ' Can't Use
        Else
        If StrComp(Cells(I, 5), SessionUsing) = 0 Or StrComp(Cells(I, 5), Unlocking) = 0 Then
            ' Red Color
            Cells(I, 4).Interior.ColorIndex = 3
        End If
    End If
    Next
    
End Sub

