Excel File => <a href="https://github.com/ShubhamDAI/VBA---Project/blob/main/SUBMIT%2CRESET%2CUPDATE%2CCANCEL%2CSEARCH%2CDELETE%20Form.xlsm"> Click Me </a>


## VBA Code

### SUBMIT
```vb
Private Sub CommandButton1_Click()

Dim lr As Long
lr = Range("A1048576").End(xlUp).Row + 1

Cells(lr, 1) = TextBox2.Value
Cells(lr, 2) = TextBox3.Value
Cells(lr, 3) = ComboBox1.Value
Cells(lr, 4) = ComboBox2.Value
Cells(lr, 5) = TextBox4.Value

TextBox2.Value = ""
TextBox3.Value = ""
ComboBox1.Value = ""
ComboBox2.Value = ""
TextBox4.Value = ""

Call refresh_data

End Sub
```

### RESET
```vb
Private Sub CommandButton2_Click()
Unload Me
UserForm1.Show
End Sub
```

### CANCEL
```vb
Private Sub CommandButton4_Click()
Unload Me
End Sub
```

### FORM ACTIVATE / DROPDOWN
```vb
Private Sub UserForm_Activate()

ComboBox1.List = Array("Manager", "Accountant", "Assistant", "Worker")
ComboBox2.List = Array("Admin", "HR", "IT", "Finance")

Range("a1") = "Emp ID"
Range("b1") = "Emp Name"
Range("c1") = "Emp Designation"
Range("d1") = "Emp Department"
Range("e1") = "Salary"

Call refresh_data
End Sub
```

### FETCHING DATA TO FORM ON DOUBLE CLICK OF LISTBOX
```vb
Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

TextBox5.Value = ListBox2.Column(0)
TextBox2.Value = ListBox2.Column(0)
TextBox3.Value = ListBox2.Column(1)
ComboBox1.Value = ListBox2.Column(2)
ComboBox2.Value = ListBox2.Column(3)
TextBox4.Value = ListBox2.Column(4)

End Sub
```

### UPDATE
```vb
Private Sub CommandButton3_Click()
Dim lr As Long
lr = Range("A1048576").End(xlUp).Row
Dim y As Long

For y = 2 To lr
    If TextBox5.Text = Cells(y, 1) Then
        Cells(y, 1) = TextBox2.Value
        Cells(y, 2) = TextBox3.Value
        Cells(y, 3) = ComboBox1.Value
        Cells(y, 4) = ComboBox2.Value
        Cells(y, 5) = TextBox4.Value
    End If
Next y

Call refresh_data
End Sub
```

### FETCHING DATA INTO LISTBOX
```vb
Sub refresh_data()

Dim lr As Long
lr = Range("A1048576").End(xlUp).Row

If lr = 1 Then lr = 2

ListBox2.ColumnCount = 5
ListBox2.ColumnHeads = True
ListBox2.RowSource = "A2:E" & lr

End Sub
```

### SEARCH
```vb
Private Sub CommandButton5_Click()
Dim lr As Long
lr = Range("A1048576").End(xlUp).Row

Dim y As Long

For y = 2 To lr
    If TextBox5.Text = Cells(y, 1) Then
        TextBox2.Value = Cells(y, 1)
        TextBox3.Value = Cells(y, 2)
        ComboBox1.Value = Cells(y, 3)
        ComboBox2.Value = Cells(y, 4)
        TextBox4.Value = Cells(y, 5)
    End If
Next y
End Sub
```

### DELETE
```vb
Private Sub CommandButton6_Click()
Dim lr As Long
lr = Range("A1048576").End(xlUp).Row
Dim y As Long

For y = 2 To lr
    If TextBox5.Text = Cells(y, 1) Then
        Rows(y).Delete
    End If
Next y

Call refresh_data
End Sub
```
