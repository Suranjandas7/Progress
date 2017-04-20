Attribute VB_Name = "DuR"
'Initial Variables
Dim LastBal As Double
Dim DateNow As String
Sub SubTAccount(txncode, amount, txndeg)
Sheets("control").Select
DateNow = Range("F1").Value

'Selects the printing sheet
    Sheets("entries").Select
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row + 1
'Different labels for different types of transactions
    If txncode = 1 Then
        Call AsstoSubT
        Range("B" & lastrow - 1).Value = "ATM"
        Range("D" & lastrow - 1).Value = amount
        Range("F" & lastrow - 1).Value = LastBal - amount
        Call degreeadder(txndeg)
    End If
    
    If txncode = 2 Then
        Call AsstoSubT
        Range("B" & lastrow - 1).Value = "POS"
        Range("D" & lastrow - 1).Value = amount
        Range("F" & lastrow - 1).Value = LastBal - amount
        Call degreeadder(txndeg)
    End If
    
    If txncode = 3 Then
        Call AsstoSubT
        Range("B" & lastrow - 1).Value = "Phone"
        Range("D" & lastrow - 1).Value = amount
        Range("F" & lastrow - 1).Value = LastBal - amount
        Call degreeadder(txndeg)
    End If
    
    If txncode = 4 Then
        Call AsstoSubT
        Range("B" & lastrow - 1).Value = "Service Charge"
        Range("D" & lastrow - 1).Value = amount
        Range("F" & lastrow - 1).Value = LastBal - amount
        Call degreeadder(txndeg)
    End If
    
    If txncode = 5 Then
        Call AsstoSubT
        Range("B" & lastrow - 1).Value = "Cash In"
        Range("E" & lastrow - 1).Value = amount
        Range("F" & lastrow - 1).Value = LastBal + amount
        Call degreeadder(txndeg)
    End If
    
    Call TAccountMaker(amount, txncode)
    Call TallyAdjuster(amount, txncode, txndeg)
End Sub
Sub AsstoSubT()
    lastrow = Cells(Rows.Count, "F").End(xlUp).Row
    LastBal = Range("F" & lastrow).Value
    
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    If UserForm1.CheckBox2.Value = True Then
        Range("A" & lastrow).Value = UserForm1.TextBox5.Value
    Else
        Range("A" & lastrow).Value = Range("Control!F1").Value
    End If
    If txncode < 5 Then
        Range("E" & lastrow).Value = " "
        Else
            Range("D" & lastrow).Value = " "
    End If
End Sub
Sub degreeadder(txndeg)
    If txndeg = 0 Then
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row + 1
        Range("A" & lastrow).Value = "Progress(c)"
    End If
    
    If txndeg = 1 Then
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        Range("C" & lastrow + 1).Value = UserForm1.ComboBox1.Value + ":" + UserForm1.TextBox2.Value
        Range("A" & lastrow + 2).Value = "Progress(c)"
    End If
    
    If txndeg = 2 Then
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        Range("C" & lastrow + 1).Value = UserForm1.ComboBox1.Value + ":" + UserForm1.TextBox2.Value
        Range("C" & lastrow + 2).Value = UserForm1.ComboBox2.Value + ":" + UserForm1.TextBox3.Value
        Range("A" & lastrow + 3).Value = "Progress(c)"
    End If
    
    If txndeg = 3 Then
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        Range("C" & lastrow + 1).Value = UserForm1.ComboBox1.Value + ":" + UserForm1.TextBox2.Value
        Range("C" & lastrow + 2).Value = UserForm1.ComboBox2.Value + ":" + UserForm1.TextBox3.Value
        Range("C" & lastrow + 3).Value = UserForm1.ComboBox3.Value + ":" + UserForm1.TextBox4.Value
        Range("A" & lastrow + 4).Value = "Progress(c)"
    End If
End Sub
Sub balanceshift()
    lastrow = Cells(Rows.Count, "C").End(xlUp).Row
    Range("C" & lastrow + 1).Value = Range("C" & lastrow)
    Range("C" & lastrow).Value = " "
    Range("F" & lastrow + 1).Value = Range("C" & lastrow)
    Range("F" & lastrow).Value = " "
    
    lastrow = Cells(Rows.Count, "E").End(xlUp).Row
    Range("E" & lastrow + 1).Value = Range("E" & lastrow)
    Range("E" & lastrow).Value = " "
End Sub
Sub TAccountMaker(amount, txncode)
    Sheets("T Account").Select
    
    If txncode < 5 Then
       lastrow = Cells(Rows.Count, "D").End(xlUp).Row + 1
        If UserForm1.CheckBox2.Value = True Then
            Range("D" & lastrow).Value = UserForm1.TextBox5.Value
        Else
            Range("D" & lastrow).Value = DateNow
        End If
       If txncode = 1 Then
        Range("E" & lastrow).Value = "ATM"
       End If
       If txncode = 2 Then
        Range("E" & lastrow).Value = "POS"
       End If
       If txncode = 3 Then
        Range("E" & lastrow).Value = "Phone"
       End If
       If txncode = 4 Then
        Range("E" & lastrow).Value = "Service Charge"
       End If
       Range("F" & lastrow).Value = amount
    Else
       lastrow = Cells(Rows.Count, "A").End(xlUp).Row + 1
       Range("A" & lastrow).Value = DateNow
       Range("B" & lastrow).Value = UserForm1.ComboBox1.Value
       Range("C" & lastrow).Value = amount
    End If
End Sub
Sub TallyAdjuster(amount, txncode, txndeg)
Sheets("Tally").Select
Dim total As Double
Dim pointerend As Integer

pointerend = Cells(Rows.Count, "I").End(xlUp).Row

'Atm
    If txncode = 1 Then
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row + 1
        Range("A" & lastrow).Value = amount
        total = Range("F2").Value
        Range("F2").Value = total + amount
'change here
        For i = 2 To 10
            If Range("I" & i).Value = UserForm1.ComboBox1.Value Then
                total = Range("K" & i).Value
                Range("K" & i).Value = total + amount
            End If
            If Range("I" & i).Value = UserForm1.ComboBox2.Value Then
                total = Range("K" & i).Value
                Range("K" & i).Value = total + amount
            End If
            If Range("I" & i).Value = UserForm1.ComboBox3.Value Then
                total = Range("K" & i).Value
                Range("K" & i).Value = total + amount
            End If
        Next i
    End If
    
'POS
    If txncode = 2 Then
        lastrow = Cells(Rows.Count, "C").End(xlUp).Row + 1
        Range("C" & lastrow).Value = amount
        total = Range("F4").Value
        Range("F4").Value = total + amount
'change here
        For i = 13 To 32
            If Range("I" & i).Value = UserForm1.ComboBox1.Value Then
                total = Range("K" & i).Value
                Range("K" & i).Value = total + amount
            End If
            If Range("I" & i).Value = UserForm1.ComboBox2.Value Then
                total = Range("K" & i).Value
                Range("K" & i).Value = total + amount
            End If
            If Range("I" & i).Value = UserForm1.ComboBox3.Value Then
                total = Range("K" & i).Value
                Range("K" & i).Value = total + amount
            End If
        Next i
    End If
    
'Phone
    If txncode = 3 Then
        lastrow = Cells(Rows.Count, "B").End(xlUp).Row + 1
        Range("B" & lastrow).Value = amount
        total = Range("F3").Value
        Range("F3").Value = total + amount
    End If
    
'Service Charges
    If txncode = 4 Then
        lastrow = Cells(Rows.Count, "D").End(xlUp).Row + 1
        Range("D" & lastrow).Value = amount
        total = Range("F5").Value
        Range("F5").Value = total + amount
    End If

'Cash in
    If txncode = 5 Then
        For i = 35 To 45
            If Range("I" & i).Value = UserForm1.ComboBox1.Value Then
                total = Range("K" & i).Value
                Range("K" & i).Value = total + amount
            End If
            If Range("I" & i).Value = UserForm1.ComboBox2.Value Then
                total = Range("K" & i).Value
                Range("K" & i).Value = total + amount
            End If
            If Range("I" & i).Value = UserForm1.ComboBox3.Value Then
                total = Range("K" & i).Value
                Range("K" & i).Value = total + amount
            End If
        Next i
    End If
End Sub
