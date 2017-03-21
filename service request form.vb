Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
       Option Compare Database

    Private Sub btnAdd_Click()
        'two options when you click on ADD
        If Me.txtSRnum.Tag & "" = "" Then
            'add data
            CurrentDb.Execute "INSERT INTO ServiceRequest(ServiceRequestNum, CategoryNum, CustomerNum, ServiceDate, Description, Status, EstHours, SpentHours, NextServiceDate, ProductNum, QuotedPrice, NumOrdered) " & _
                " VALUES(" & Me.txtSRnum & ",'" & Me.txtCatNum & "','" & Me.txtCnum & "','" & _
                Me.txtSdate & "','" & Me.txtDesc & "','" & Me.txtStatus & "','" & Me.txtEhours & "','" & _
                Me.txtShours & "','" & Me.txtNservice & "','" & Me.txtPnum & "','" & Me.txtQprice & "','" & Me.txtNumOr & "') "
        Else
            CurrentDb.Execute "UPDATE ServiceRequest " & _
                " SET ServiceRequestNum=" & Me.txtSRnum & _
                ", CategoryNum='" & Me.txtCatNum & "'" & _
                ", CustomerNum='" & Me.txtCnum & "'" & _
                ", ServiceDate='" & Me.txtSdate & "'" & _
                ", Description='" & Me.txtDesc & "'" & _
                ", Status='" & Me.txtStatus & "'" & _
                ", EstHours='" & Me.txtEhours & "'" & _
                ", SpentHours='" & Me.txtShours & "'" & _
                ", NextServiceDate='" & Me.txtNservice & "'" & _
                ", ProductNum='" & Me.txtPnum & "'" & _
                ", QuotedPrice='" & Me.txtQprice & "'" & _
                ", NumOrdered='" & Me.txtNumOr & "'" & _
                " WHERE ServiceRequestNum=" & Me.txtSRnum.Tag

        End If

        'clear form
        btnClear_Click()
        'refresh data
        ServiceRequest_SubForm.Form.Requery()
    End Sub

    Private Sub btnClear_Click()
        Me.txtSRnum = ""
        Me.txtCatNum = ""
        Me.txtCnum = ""
        Me.txtSdate = ""
        Me.txtDesc = ""
        Me.txtStatus = ""
        Me.txtEhours = ""
        Me.txtShours = ""
        Me.txtNservice = ""
        Me.txtPnum = ""
        Me.txtQprice = ""
        Me.txtNumOr = ""

        Me.txtSRnum.SetFocus()
        Me.btnEdit.Enabled = True
        Me.btnAdd.Caption = "Add"
        Me.txtSRnum.Tag = ""
    End Sub

    Private Sub btnClose_Click()
        DoCmd.Close()
    End Sub

    Private Sub btnDelete_Click()
        If Not (Me.ServiceRequest_SubForm.Form.Recordset.EOF And Me.ServiceRequest_SubForm.Form.Recordset.EOF) Then
            If MsgBox("Are you sure you want to delete?", vbYesNo) = vbYes Then
                CurrentDb.Execute "DELETE FROM ServiceRequest " & _
                    " WHERE ServiceRequestNum=" & Me.ServiceRequest_SubForm.Form.Recordset.Fields("ServiceRequestNum")
                Me.ServiceRequest_SubForm.Form.Requery()
            End If
        End If
    End Sub

    Private Sub btnEdit_Click()
        If Not (Me.ServiceRequest_SubForm.Form.Recordset.EOF And Me.ServiceRequest_SubForm.Form.Recordset.BOF) Then
            With Me.ServiceRequest_SubForm.Form.Recordset
                Me.txtSRnum = .Fields("ServiceRequestNum")
                Me.txtCatNum = .Fields("CategoryNum")
                Me.txtCnum = .Fields("CustomerNum")
                Me.txtSdate = .Fields("ServiceDate")
                Me.txtDesc = .Fields("Description")
                Me.txtStatus = .Fields("Status")
                Me.txtEhours = .Fields("EstHours")
                Me.txtShours = .Fields("SpentHours")
                Me.txtNservice = .Fields("NextServiceDate")
                Me.txtPnum = .Fields("ProductNum")
                Me.txtQprice = .Fields("QuotedPrice")
                Me.txtNumOr = .Fields("NumOrdered")
                'store customer number in Tag of txtCnum in case id is modified
                Me.txtSRnum.Tag = .Fields("ServiceRequestNum")
                'change caption of add to update
                Me.btnAdd.Caption = "Update"
                'disable button edit
                Me.btnEdit.Enabled = False'
            End With
        End If
    End Sub
