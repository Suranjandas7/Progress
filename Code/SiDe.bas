Attribute VB_Name = "SiDe"
Sub ReallocateInvestment()

End Sub
Sub OFP()

End Sub
Sub ProgReport()

End Sub
Sub viewaexpenses()
    Sheets("MBS").Select
    lastrow = Cells(Rows.Count, "C").End(xlUp).Row
    Dim NumAc As Integer
    NumAc = lastrow - 12
    
    Dim Expected(0 To 20) As Double
    Dim Actual(0 To 20) As Double
    Dim labels(0 To 20) As String
    
    For i = 0 To 20
        Expected(i) = Range("C" & i + 12).Value
        Actual(i) = Range("B" & i + 12).Value
        labels(i) = Range("A" & i + 12).Value
    Next i
    
    For i = 0 To 20
        If Expected(i) > Actual(i) Then
            MsgBox "You can spend INR." & Expected(i) - Actual(i) & " on " & labels(i)
        End If
    Next i
    
    MsgBox "Thanks! Redirecting to dashboard..."
    Sheets("Dash").Select
End Sub
Sub vieweincomes()
    Sheets("MBS").Select
    lastrow = Cells(Rows.Count, "F").End(xlUp).Row
    Dim NumAc As Integer
    NumAc = lastrow - 12
    
    Dim Expected(0 To 20) As Double
    Dim Actual(0 To 20) As Double
    Dim labels(0 To 20) As String
    
    For i = 0 To 20
        Expected(i) = Range("F" & i + 12).Value
        Actual(i) = Range("E" & i + 12).Value
        labels(i) = Range("D" & i + 12).Value
    Next i
    
    For i = 0 To 20
        If Actual(i) < Expected(i) Then
            MsgBox "You will generate INR." & Expected(i) - Actual(i) & " from " & labels(i)
        End If
    Next i
    MsgBox "Thanks! Redirecting to dashboard..."
    Sheets("Dash").Select
End Sub
Sub addaccount()

End Sub
Sub SpendPattern()
    Sheets("MBS").Select
    Dim Expenses(1 To 100) As Double
    Dim Temp As String
    Dim Temp2 As Double
    lastrow = Cells(Rows.Count, "B").End(xlUp).Row
    Dim total As Double
    total = 0
    Dim UserAmt As Double
    
    UserAmt = InputBox("Enter the amount of money:")

    For i = 12 To lastrow
        Temp = Range("B" & i).Value
        Temp2 = CDbl(Temp)
        Expenses(i - 11) = Temp2
    Next i
    
    Dim PerCent(1 To 100) As Double
    Dim labels(1 To 100) As String
    
    For i = 1 To lastrow - 11
        labels(i) = Range("A" & 11 + i)
        total = total + Expenses(i)
    Next i
    
    For i = 1 To lastrow - 11
        PerCent(i) = Expenses(i) / total
        PerCent(i) = Round(PerCent(i), 2)
    Next i
    
    Dim ShowUAMT As String
    ShowUAMT = CStr(UserAmt)

'UserForm adj

    UserForm2.Label1.Caption = ShowUAMT
    
    UserForm2.AC1.Caption = labels(1)
    UserForm2.P1.Caption = PerCent(1)
    UserForm2.AM1.Caption = PerCent(1) * UserAmt
    
    UserForm2.AC2.Caption = labels(2)
    UserForm2.P2.Caption = PerCent(2)
    UserForm2.AM2.Caption = PerCent(2) * UserAmt
    
    UserForm2.AC3.Caption = labels(3)
    UserForm2.P3.Caption = PerCent(3)
    UserForm2.AM3.Caption = PerCent(3) * UserAmt
    
    UserForm2.AC4.Caption = labels(4)
    UserForm2.P4.Caption = PerCent(4)
    UserForm2.AM4.Caption = PerCent(4) * UserAmt
    
    UserForm2.AC5.Caption = labels(5)
    UserForm2.AM5.Caption = PerCent(5) * UserAmt
    UserForm2.P5.Caption = PerCent(5)
    
    UserForm2.AC6.Caption = labels(6)
    UserForm2.AM6.Caption = PerCent(6) * UserAmt
    UserForm2.P6.Caption = PerCent(6)
    
    UserForm2.AC7.Caption = labels(7)
    UserForm2.AM7.Caption = PerCent(7) * UserAmt
    UserForm2.P7.Caption = PerCent(7)
    
    UserForm2.AC8.Caption = labels(8)
    UserForm2.AM8.Caption = PerCent(8) * UserAmt
    UserForm2.P8.Caption = PerCent(8)
    
    UserForm2.AC9.Caption = labels(9)
    UserForm2.AM9.Caption = PerCent(9) * UserAmt
    UserForm2.P9.Caption = PerCent(9)
    
UserForm2.Show
Sheets("Dash").Select
End Sub
