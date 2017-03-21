Option Compare Database

    Private Sub btnClear_Click()
        Me.txtPnum = ""
        Me.txtDesc = ""
        Me.txtOhand = ""
        Me.txtCat = ""
        Me.txtPrice = ""

        Me.txtPnum.SetFocus()
        Me.btnEdit.Enabled = True
        Me.btnAdd.Caption = "Add"
        Me.txtPnum.Tag = ""
    End Sub

    Private Sub btnClose_Click()
        DoCmd.Close()
    End Sub

    Private Sub btnDelete_Click()
        If Not (Me.Product_SubForm.Form.Recordset.EOF And Me.Product_SubForm.Form.Recordset.EOF) Then
            If MsgBox("Are you sure you want to delete?", vbYesNo) = vbYes Then
                CurrentDb.Execute "DELETE FROM Product " & _
                    " WHERE ProductNum=" & Me.Product_SubForm.Form.Recordset.Fields("ProductNum")
                Me.Product_SubForm.Form.Requery()
            End If
        End If
    End Sub

    Private Sub btnEdit_Click()
        If Not (Me.Product_SubForm.Form.Recordset.EOF And Me.Product_SubForm.Form.Recordset.BOF) Then
            With Me.Product_SubForm.Form.Recordset
                Me.txtPnum = .Fields("ProductNum")
                Me.txtDesc = .Fields("Description")
                Me.txtOhand = .Fields("OnHand")
                Me.txtCat = .Fields("Category")
                Me.txtPrice = .Fields("Price")
                'store customer number in Tag of txtCnum in case id is modified
                Me.txtPnum.Tag = .Fields("ProductNum")
                'change caption of add to update
                Me.btnAdd.Caption = "Update"
                'disable button edit
                Me.btnEdit.Enabled = False
            End With
        End If
    End Sub

    Private Sub btnAdd_Click()
        'two options when you click on ADD
        If Me.txtPnum.Tag & "" = "" Then
            'add data
            CurrentDb.Execute "INSERT INTO Product(ProductNum, Description, OnHand, Category, Price) " & _
                " VALUES(" & Me.txtPnum & ",'" & Me.txtDesc & "','" & Me.txtOhand & "','" & Me.txtCat & "','" & Me.txtPrice & "') "
        Else
            CurrentDb.Execute "UPDATE Product " & _
                " SET ProductNum=" & Me.txtPnum & _
                ", Description='" & Me.txtDesc & "'" & _
                ", OnHand='" & Me.txtOhand & "'" & _
                ", Category='" & Me.txtCat & "'" & _
                ", Price='" & Me.txtPrice & "'" & _
                " WHERE ProductNum=" & Me.txtPnum.Tag

        End If

        'clear form
        btnClear_Click()
        'refresh data
        Product_SubForm.Form.Requery()
    End Sub
