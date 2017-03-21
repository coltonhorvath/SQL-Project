/*Written in Visual Basic*/

Option Compare Database

    Private Sub btnAdd_Click()
        'two options when you click on ADD
        If Me.txtTnum.Tag & "" = "" Then
            'add data
            CurrentDb.Execute "INSERT INTO Techs(TechNum, LastName, FirstName, Street, City, State, PostalCode, Bonus, Rate) " & _
                " VALUES(" & Me.txtTnum & ",'" & Me.txtLname & "','" & Me.txtFname & "','" & Me.txtStreet & "','" & Me.txtCity & "','" & Me.txtState & "','" & Me.txtPostal & "','" & Me.txtBonus & "','" & Me.txtRate & "') "
        Else
            CurrentDb.Execute "UPDATE Techs " & _
                " SET TechNum=" & Me.txtTnum & _
                ", LastName='" & Me.txtLname & "'" & _
                ", FirstName='" & Me.txtFname & "'" & _
                ", Street='" & Me.txtStreet & "'" & _
                ", City='" & Me.txtCity & "'" & _
                ", State='" & Me.txtState & "'" & _
                ", PostalCode='" & Me.txtPostal & "'" & _
                ", Bonus='" & Me.txtBonus & "'" & _
                ", Rate='" & Me.txtRate & "'" & _
                " WHERE TechNum=" & Me.txtTnum.Tag

        End If

        'clear form
        btnClear_Click()
        'refresh data
        Techs_SubForm.Form.Requery()
    End Sub

    Private Sub btnClear_Click()
        Me.txtTnum = ""
        Me.txtLname = ""
        Me.txtFname = ""
        Me.txtStreet = ""
        Me.txtCity = ""
        Me.txtState = ""
        Me.txtPostal = ""
        Me.txtBonus = ""
        Me.txtRate = ""

        Me.txtTnum.SetFocus()
        Me.btnEdit.Enabled = True
        Me.btnAdd.Caption = "Add"
        Me.txtTnum.Tag = ""
    End Sub

    Private Sub btnClose_Click()
        DoCmd.Close()
    End Sub

    Private Sub btnDelete_Click()
        If Not (Me.Techs_SubForm.Form.Recordset.EOF And Me.Techs_SubForm.Form.Recordset.EOF) Then
            If MsgBox("Are you sure you want to delete?", vbYesNo) = vbYes Then
                CurrentDb.Execute "DELETE FROM Techs " & _
                    " WHERE TechNum=" & Me.Techs_SubForm.Form.Recordset.Fields("TechNum")
                Me.Techs_SubForm.Form.Requery()
            End If
        End If
    End Sub

    Private Sub btnEdit_Click()
        If Not (Me.Techs_SubForm.Form.Recordset.EOF And Me.Techs_SubForm.Form.Recordset.BOF) Then
            With Me.Techs_SubForm.Form.Recordset
                Me.txtTnum = .Fields("TechNum")
                Me.txtLname = .Fields("LastName")
                Me.txtFname = .Fields("FirstName")
                Me.txtStreet = .Fields("Street")
                Me.txtCity = .Fields("City")
                Me.txtState = .Fields("State")
                Me.txtPostal = .Fields("PostalCode")
                Me.txtBonus = .Fields("Bonus")
                Me.txtRate = .Fields("Rate")
                'store customer number in Tag of txtCnum in case id is modified
                Me.txtTnum.Tag = .Fields("TechNum")
                'change caption of add to update
                Me.btnAdd.Caption = "Update"
                'disable button edit
                Me.btnEdit.Enabled = False
            End With
        End If

    End Sub
