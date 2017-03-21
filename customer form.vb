Option Compare Database

    Private Sub btnAdd_Click()
        'two options when you click on ADD
        If Me.txtCnum.Tag & "" = "" Then
            'add data
            CurrentDb.Execute "INSERT INTO Customer(CustomerNum, CustomerName, Street, City, State, PostalCode, Balance, TechNum) " & _
                " VALUES(" & Me.txtCnum & ",'" & Me.txtCname & "','" & Me.txtStreet & "','" & Me.txtCity & "','" & Me.txtState & "','" & Me.txtPostal & "','" & Me.txtBalance & "','" & Me.txtTnum & "') "
        Else
            CurrentDb.Execute "UPDATE Customer " & _
                " SET CustomerNum=" & Me.txtCnum & _
                ", CustomerName='" & Me.txtCname & "'" & _
                ", Street='" & Me.txtStreet & "'" & _
                ", City='" & Me.txtCity & "'" & _
                ", State='" & Me.txtState & "'" & _
                ", PostalCode='" & Me.txtPostal & "'" & _
                ", Balance='" & Me.txtBalance & "'" & _
                ", TechNum='" & Me.txtTnum & "'" & _
                " WHERE CustomerNum=" & Me.txtCnum.Tag

        End If

        'clear form
        btnClear_Click()
        'refresh data
        Customer_Subform.Form.Requery()
    End Sub

    Private Sub btnClear_Click()
        Me.txtCnum = ""
        Me.txtCname = ""
        Me.txtStreet = ""
        Me.txtCity = ""
        Me.txtState = ""
        Me.txtPostal = ""
        Me.txtBalance = ""
        Me.txtTnum = ""

        Me.txtCnum.SetFocus()
        Me.btnEdit.Enabled = True
        Me.btnAdd.Caption = "Add"
        Me.txtCnum.Tag = ""


    End Sub

    Private Sub btnClose_Click()
        DoCmd.Close()
    End Sub

    Private Sub btnDelete_Click()
        If Not (Me.Customer_Subform.Form.Recordset.EOF And Me.Customer_Subform.Form.Recordset.EOF) Then
            If MsgBox("Are you sure you want to delete?", vbYesNo) = vbYes Then
                CurrentDb.Execute "DELETE FROM Customer " & _
                    " WHERE CustomerNum=" & Me.Customer_Subform.Form.Recordset.Fields("CustomerNum")
                Me.Customer_Subform.Form.Requery()
            End If
        End If
    End Sub

    Private Sub btnEdit_Click()
        If Not (Me.Customer_Subform.Form.Recordset.EOF And Me.Customer_Subform.Form.Recordset.BOF) Then
            With Me.Customer_Subform.Form.Recordset
                Me.txtCnum = .Fields("CustomerNum")
                Me.txtCname = .Fields("CustomerName")
                Me.txtStreet = .Fields("Street")
                Me.txtCity = .Fields("City")
                Me.txtState = .Fields("State")
                Me.txtPostal = .Fields("PostalCode")
                Me.txtBalance = .Fields("Balance")
                Me.txtTnum = .Fields("TechNum")
                'store customer number in Tag of txtCnum in case id is modified
                Me.txtCnum.Tag = .Fields("CustomerNum")
                'change caption of add to update
                Me.btnAdd.Caption = "Update"
                'disable button edit
                Me.btnEdit.Enabled = False
            End With
        End If

    End Sub
