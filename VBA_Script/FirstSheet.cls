VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub First()

Dim X As Double, RealCount As Integer, I As Integer, J As Integer, Offset As Integer
Dim Txt As String, ResultInteger As Integer, MinusCount As Integer
Dim arrList As Object
Set arrList = CreateObject("System.Collections.ArrayList")
Dim secondArrList As Object
Set secondArrList = CreateObject("System.Collections.ArrayList")
Dim UnConnected As String
Dim ResultIndex As Integer

' Directory Address
Open "\\##.#.###.#\######\######\Check####Log.txt" For Input As #1 'Any Number you want to assign

' Starting Point, Initial Count
X = 0

    ' Run the procedure while it is not at end of the file, Parameter is FileNumber
    Do While Not EOF(1)
    Line Input #1, Txt
    Cells(1, 1).Offset(X, 0) = Txt
    X = X + 1
    Loop
    
Close #1

' Count Rows, Initialized
RealCount = 0
X = 1

Do While True
    RealCount = RealCount + 1
    
    If IsEmpty(Cells(X, 1)) And IsEmpty(Cells(X + 1, 1)) And IsEmpty(Cells(X + 2, 1)) Then
          Exit Do
    End If
    X = X + 1
    
Loop

' Result Rows Count
Cells(RealCount, 1) = RealCount

' Sort Use / Don't Use & Initialized
TemporaryCount = 1
Offset = 13

' Sort results via For-Loop
For I = 3 To RealCount - 1 Step 13

CutStringUserName = Cells(Offset - 3 + I, 1)
CutStringPcName = Cells(Offset - 12 + I, 1)

' Get Time, LogOnID, MachineName, Message Each Cells
Cells(RealCount + TemporaryCount, 1) = Cells(I, 1)
Cells(RealCount + TemporaryCount, 2) = Right(CutStringUserName, 10)
Cells(RealCount + TemporaryCount, 3) = Right(CutStringPcName, 22)
Cells(RealCount + TemporaryCount, 4) = Cells(Offset - 9 + I, 1)

TemporaryCount = TemporaryCount + 1

' Use ResultInteger And Initialized
ResultInteger = RealCount + TemporaryCount
Next

MinusCount = 0
J = ResultInteger - 1

' Filtering if different Data
For I = RealCount + 1 To J
If StrComp(Cells(I, 3), Cells(I + 1, 3)) <> 0 Then
Cells(ResultInteger + 3, 4) = Cells(I - MinusCount, 4)
Cells(ResultInteger + 3, 3) = Cells(I - MinusCount, 3)
Cells(ResultInteger + 3, 2) = Cells(I - MinusCount, 2)
Cells(ResultInteger + 3, 1) = Cells(I - MinusCount, 1)
MinusCount = 0
ResultInteger = ResultInteger + 1
Else
    MinusCount = MinusCount + 1
End If

Next
          
' Last Sort
TemporaryCount = 1
For I = J + 4 To ResultInteger + 2
    ' If No Element
  If arrList.Count < 1 Then
      Cells(TemporaryCount, 8) = Cells(I, 4)
      Cells(TemporaryCount, 7) = Cells(I, 3)
      Cells(TemporaryCount, 6) = Cells(I, 2)
      Cells(TemporaryCount, 5) = Cells(I, 1)
      arrList.Add (Cells(I, 3))
  Else
  
  For K = 0 To arrList.Count - 1
    If StrComp(Cells(I, 3), arrList(K)) = 0 Then
    GoTo ContinueDo
    End If
    Next
    End If
    
  If arrList.Count > 0 Then
    arrList.Add (Cells(I, 3))
    Cells(TemporaryCount, 8) = Cells(I, 4)
    Cells(TemporaryCount, 7) = Cells(I, 3)
    Cells(TemporaryCount, 6) = Cells(I, 2)
    Cells(TemporaryCount, 5) = Cells(I, 1)
    TemporaryCount = TemporaryCount + 1
    End If

ContinueDo: Next

ResultIndex = TemporaryCount
    
' Second Files
' Directory Address
Open "#:\######\##########.txt" For Input As #2 'Any Number you want to assign File Number

' Reset Point
X = 0

    Do While Not EOF(2)
    Line Input #2, Txt
    Cells(1, 10).Offset(X, 0) = Txt
    X = X + 1
    Loop
    
Close #2


' Second Count Rows, Initialized
RealCount = 0
X = 1

Do While True
    RealCount = RealCount + 1
    
    If IsEmpty(Cells(X, 10)) And IsEmpty(Cells(X + 1, 10)) And IsEmpty(Cells(X + 2, 10)) Then
          Exit Do
    End If
    X = X + 1
    
Loop

' Result Rows Count
Cells(RealCount, 10) = RealCount

' Check Condition To Initialized
UnConnected = "##########################################"

' Sort Use / Don't Use & Initialized
TemporaryCount = 1
Offset = 21

' Sort results via For-Loop
For I = 3 To RealCount - 1 Step 21

CutStringUserName = Cells(Offset - 12 + I, 10)
CutStringPcName = Cells(Offset - 20 + I, 10)

' Get Time, LogOnID, MachineName, Message Each Cells
Cells(RealCount + TemporaryCount, 10) = Cells(I, 10)
Cells(RealCount + TemporaryCount, 11) = Right(CutStringUserName, 10)
Cells(RealCount + TemporaryCount, 12) = Right(CutStringPcName, 22)
Cells(RealCount + TemporaryCount, 13) = Cells(Offset - 17 + I, 10)

If StrComp(Cells(Offset - 17 + I, 10), UnConnected) = 0 Then
I = I + 1
End If

TemporaryCount = TemporaryCount + 1

' Use ResultInteger And Initialized
ResultInteger = RealCount + TemporaryCount
Next

MinusCount = 0
J = ResultInteger - 1

' Filtering if different Data
For I = RealCount + 1 To J
If StrComp(Cells(I, 12), Cells(I + 1, 12)) <> 0 Then
Cells(ResultInteger + 3, 10) = Cells(I - MinusCount, 10)
Cells(ResultInteger + 3, 11) = Cells(I - MinusCount, 11)
Cells(ResultInteger + 3, 12) = Cells(I - MinusCount, 12)
Cells(ResultInteger + 3, 13) = Cells(I - MinusCount, 13)
MinusCount = 0
ResultInteger = ResultInteger + 1
Else
    MinusCount = MinusCount + 1
End If

Next
     
For I = J + 4 To ResultInteger + 2
    ' If No Element
  If secondArrList.Count < 1 Then
      Cells(ResultIndex, 5) = Cells(I, 10)
      Cells(ResultIndex, 6) = Cells(I, 11)
      Cells(ResultIndex, 7) = Cells(I, 12)
      Cells(ResultIndex, 8) = Cells(I, 13)
      secondArrList.Add (Cells(I, 12))
  Else
  
  For K = 0 To secondArrList.Count - 1
    If StrComp(Cells(I, 12), secondArrList(K)) = 0 Then
    GoTo ContinueSecond
    End If
    Next
    End If
    
  If secondArrList.Count > 0 Then
    secondArrList.Add (Cells(I, 12))
    Cells(ResultIndex, 5) = Cells(I, 10)
    Cells(ResultIndex, 6) = Cells(I, 11)
    Cells(ResultIndex, 7) = Cells(I, 12)
    Cells(ResultIndex, 8) = Cells(I, 13)
    ResultIndex = ResultIndex + 1
    End If
    
    SecondLastIndex = ResultInteger + 2 + TemporaryCount
ContinueSecond:     Next
    
    
End Sub




