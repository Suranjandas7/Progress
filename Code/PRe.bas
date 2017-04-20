Attribute VB_Name = "PRe"
Dim txncode As Integer
Dim CrODr As Integer
Dim amount As Double
Dim txndeg As Integer
Sub EM()
    UserForm1.Show
End Sub
Sub onclick()

'Hides the userform
    UserForm1.Hide
    txncode = 999
    CrODr = 0

'checks the deg of txn
    If UserForm1.ComboBox3.Value = "-" Then
        If UserForm1.ComboBox2.Value = "-" Then
            If UserForm1.ComboBox1.Value = "-" Then
                txndeg = 0
                GoTo gotohere
            Else
                txndeg = 1
                GoTo gotohere
            End If
        End If
        txndeg = 2
        Else
            txndeg = 3
    End If
gotohere:

'checks the txncode of the transaction
    If UserForm1.OptionButton1.Value = True Then
    'Atm
        txncode = 1
    End If
    If UserForm1.OptionButton2.Value = True Then
    'pos
        txncode = 2
    End If
    If UserForm1.OptionButton3.Value = True Then
    'phone
        txncode = 3
    End If
    If UserForm1.OptionButton4.Value = True Then
    'service charge
        txncode = 4
    End If
    If UserForm1.OptionButton5.Value = True Then
    'cash in
        txncode = 5
    End If

'fetches form value
    amount = UserForm1.TextBox1.Value
    
'checks to debit or credit the account
    If UserForm1.CheckBox1.Value = True Then
        MsgBox ("Money will be added with this transaction.")
        CrODr = 1
    End If
    If UserForm1.CheckBox1.Value = False Then
        MsgBox ("Money will be deducted with this transaction.")
        CrODr = 0
    End If
    Call SubTAccount(txncode, amount, txndeg)
End Sub
