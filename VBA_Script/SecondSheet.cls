VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Second()
Dim lastRow As Long, FirstIndex As Integer, TemporaryInteger As Integer
Dim ArrayList As Object
Set ArrayList = CreateObject("System.Collections.ArrayList")

lastRow = Sheet1.Cells(Rows.Count, 5).End(xlUp).Row
    ' Call Sheet1 Last Information
    For I = 1 To lastRow
        Cells(I, 2) = Sheet1.Cells(I, 5)
        Cells(I, 3) = Sheet1.Cells(I, 6)
        Cells(I, 4) = Sheet1.Cells(I, 7)
        Cells(I, 5) = Sheet1.Cells(I, 8)
    Next

    ' Newest Information
    
    TemporaryInteger = 3
    FirstIndex = 1
    For I = 1 To lastRow
        For J = I + 1 To lastRow
        
        ' Compare Id
        If StrComp(Cells(I, 4), Cells(J, 4)) = 0 Then
        ' Compare Time
            If StrComp(Cells(I, 2), Cells(J, 2)) = 0 Or StrComp(Cells(I, 2), Cells(J, 2)) = 1 Then
                Cells(FirstIndex, 7) = Cells(I, 2)
                 Cells(FirstIndex, 8) = Cells(I, 3)
                  Cells(FirstIndex, 9) = Cells(I, 4)
                   Cells(FirstIndex, 10) = Cells(I, 5)
                   ArrayList.Add (Cells(FirstIndex, 9))
                    FirstIndex = FirstIndex + 1
                   GoTo ContinueNext
            Else
                    
           CheckDuplication = True
                 Cells(FirstIndex, 7) = Cells(J, 2)
                  Cells(FirstIndex, 8) = Cells(J, 3)
                   Cells(FirstIndex, 9) = Cells(J, 4)
                    Cells(FirstIndex, 10) = Cells(J, 5)
                       ArrayList.Add (Cells(FirstIndex, 9))
                     FirstIndex = FirstIndex + 1
                     GoTo ContinueNext
            End If
        End If
      
        Next
        TemporaryInteger = FirstIndex
ContinueNext:    Next
       
        ' Call Not Duplicate Data

        For I = 1 To lastRow
           For J = 0 To ArrayList.Count - 1
        
            If StrComp(Cells(I, 4), ArrayList(J)) = 0 Then
               GoTo NextIndex
               End If
            Next
            
              Cells(TemporaryInteger, 7) = Cells(I, 2)
               Cells(TemporaryInteger, 8) = Cells(I, 3)
                 Cells(TemporaryInteger, 9) = Cells(I, 4)
                  Cells(TemporaryInteger, 10) = Cells(I, 5)
                  TemporaryInteger = TemporaryInteger + 1
NextIndex: Next
        
   
End Sub

