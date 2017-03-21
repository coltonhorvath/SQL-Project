Option Compare Database

    Private Sub btnAdd_Click()
        'two options when you click on ADD
        If Me.txtCatNum.Tag & "" = "" Then
            'add data
            CurrentDb.Execute "INSERT INTO ServiceCategory(CategoryNum, CategoryDescription) " & _
                " VALUES(" & Me.txtCatNum & ",'" & Me.txtCatDesc & "') "
        Else
            CurrentDb.Execute "UPDATE ServiceCategory " & _
                " SET CategoryNum=" & Me.txtCatNum & _
                ", CategoryDescription='" & Me.txtCatDesc & "'" & _
                " WHERE CategoryNum=" & Me.txtCatNum.Tag

        End If

        'clear form
        btnClear_Click()
        'refresh data
        ServiceCategory_SubForm.Form.Requery()
    End Sub

    Private Sub btnClear_Click()
        Me.txtCatNum = ""
        Me.txtCatDesc = ""

        Me.txtCatNum.SetFocus()
        Me.btnEdit.Enabled = True
        Me.btnAdd.Caption = "Add"
        Me.txtCatNum.Tag = ""
    End Sub

    Private Sub btnClose_Click()
        DoCmd.Close()
    End Sub

    Private Sub btnDelete_Click()
        If Not (Me.ServiceCategory_SubForm.Form.Recordset.EOF And Me.ServiceCategory_SubForm.Form.Recordset.EOF) Then
            If MsgBox("Are you sure you want to delete?", vbYesNo) = vbYes Then
                CurrentDb.Execute "DELETE FROM ServiceCategory " & _
                    " WHERE CategoryNum=" & Me.ServiceCategory_SubForm.Form.Recordset.Fields("CategoryNum")
                Me.ServiceCategory_SubForm.Form.Requery()
            End If
        End If
    End Sub

    Private Sub btnEdit_Click()
        If Not (Me.ServiceCategory_SubForm.Form.Recordset.EOF And Me.ServiceCategory_SubForm.Form.Recordset.BOF) Then
            With Me.ServiceCategory_SubForm.Form.Recordset
                Me.txtCatNum = .Fields("CategoryNum")
                Me.txtCatDesc = .Fields("CategoryDescription")
                'store customer number in Tag of txtCnum in case id is modified
                Me.txtCatNum.Tag = .Fields("CategoryNum")
                'change caption of add to update
                Me.btnAdd.Caption = "Update"
                'disable button edit
                Me.btnEdit.Enabled = False
            End With
        End If
    End Sub