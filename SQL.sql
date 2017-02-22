Techs Form

Option Compare Database

Private Sub btnAdd_Click()

&#39;two options when you click on ADD

If Me.txtTnum.Tag &amp; &quot;&quot; = &quot;&quot; Then

&#39;add data

CurrentDb.Execute &quot;INSERT INTO Techs(TechNum, LastName, FirstName,

Street, City, State, PostalCode, Bonus, Rate) &quot; &amp; _

&quot; VALUES(&quot; &amp; Me.txtTnum &amp; &quot;,&#39;&quot; &amp; Me.txtLname &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtFname

&amp; &quot;&#39;,&#39;&quot; &amp; Me.txtStreet &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtCity &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtState &amp; &quot;&#39;,&#39;&quot; &amp;

Me.txtPostal &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtBonus &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtRate &amp; &quot;&#39;) &quot;

Else

CurrentDb.Execute &quot;UPDATE Techs &quot; &amp; _

&quot; SET TechNum=&quot; &amp; Me.txtTnum &amp; _

&quot;, LastName=&#39;&quot; &amp; Me.txtLname &amp; &quot;&#39;&quot; &amp; _

&quot;, FirstName=&#39;&quot; &amp; Me.txtFname &amp; &quot;&#39;&quot; &amp; _

&quot;, Street=&#39;&quot; &amp; Me.txtStreet &amp; &quot;&#39;&quot; &amp; _

&quot;, City=&#39;&quot; &amp; Me.txtCity &amp; &quot;&#39;&quot; &amp; _

&quot;, State=&#39;&quot; &amp; Me.txtState &amp; &quot;&#39;&quot; &amp; _

&quot;, PostalCode=&#39;&quot; &amp; Me.txtPostal &amp; &quot;&#39;&quot; &amp; _

&quot;, Bonus=&#39;&quot; &amp; Me.txtBonus &amp; &quot;&#39;&quot; &amp; _

&quot;, Rate=&#39;&quot; &amp; Me.txtRate &amp; &quot;&#39;&quot; &amp; _

&quot; WHERE TechNum=&quot; &amp; Me.txtTnum.Tag

End If

&#39;clear form

btnClear_Click()

&#39;refresh data

Techs_SubForm.Form.Requery()

End Sub

Private Sub btnClear_Click()

Me.txtTnum = &quot;&quot;

Me.txtLname = &quot;&quot;

Me.txtFname = &quot;&quot;

Me.txtStreet = &quot;&quot;

Me.txtCity = &quot;&quot;

Me.txtState = &quot;&quot;

Me.txtPostal = &quot;&quot;

Me.txtBonus = &quot;&quot;

Me.txtRate = &quot;&quot;

Me.txtTnum.SetFocus()

Me.btnEdit.Enabled = True

Me.btnAdd.Caption = &quot;Add&quot;

Me.txtTnum.Tag = &quot;&quot;

End Sub

Private Sub btnClose_Click()

DoCmd.Close()

End Sub

11

Private Sub btnDelete_Click()

If Not (Me.Techs_SubForm.Form.Recordset.EOF And

Me.Techs_SubForm.Form.Recordset.EOF) Then

If MsgBox(&quot;Are you sure you want to delete?&quot;, vbYesNo) = vbYes Then

CurrentDb.Execute &quot;DELETE FROM Techs &quot; &amp; _

&quot; WHERE TechNum=&quot; &amp;

Me.Techs_SubForm.Form.Recordset.Fields(&quot;TechNum&quot;)

Me.Techs_SubForm.Form.Requery()

End If

End If

End Sub

Private Sub btnEdit_Click()

If Not (Me.Techs_SubForm.Form.Recordset.EOF And

Me.Techs_SubForm.Form.Recordset.BOF) Then

With Me.Techs_SubForm.Form.Recordset

Me.txtTnum = .Fields(&quot;TechNum&quot;)

Me.txtLname = .Fields(&quot;LastName&quot;)

Me.txtFname = .Fields(&quot;FirstName&quot;)

Me.txtStreet = .Fields(&quot;Street&quot;)

Me.txtCity = .Fields(&quot;City&quot;)

Me.txtState = .Fields(&quot;State&quot;)

Me.txtPostal = .Fields(&quot;PostalCode&quot;)

Me.txtBonus = .Fields(&quot;Bonus&quot;)

Me.txtRate = .Fields(&quot;Rate&quot;)

&#39;store customer number in Tag of txtCnum in case id is modified

Me.txtTnum.Tag = .Fields(&quot;TechNum&quot;)

&#39;change caption of add to update

Me.btnAdd.Caption = &quot;Update&quot;

&#39;disable button edit

Me.btnEdit.Enabled = False

End With

End If

End Sub

Service Request Form

Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

Option Compare Database

Private Sub btnAdd_Click()

&#39;two options when you click on ADD

If Me.txtSRnum.Tag &amp; &quot;&quot; = &quot;&quot; Then

&#39;add data

CurrentDb.Execute &quot;INSERT INTO ServiceRequest(ServiceRequestNum,

CategoryNum, CustomerNum, ServiceDate, Description, Status, EstHours, SpentHours,

NextServiceDate, ProductNum, QuotedPrice, NumOrdered) &quot; &amp; _

&quot; VALUES(&quot; &amp; Me.txtSRnum &amp; &quot;,&#39;&quot; &amp; Me.txtCatNum &amp; &quot;&#39;,&#39;&quot; &amp;

Me.txtCnum &amp; &quot;&#39;,&#39;&quot; &amp; _

Me.txtSdate &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtDesc &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtStatus &amp; &quot;&#39;,&#39;&quot; &amp;

Me.txtEhours &amp; &quot;&#39;,&#39;&quot; &amp; _

Me.txtShours &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtNservice &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtPnum &amp; &quot;&#39;,&#39;&quot;

&amp; Me.txtQprice &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtNumOr &amp; &quot;&#39;) &quot;

Else

CurrentDb.Execute &quot;UPDATE ServiceRequest &quot; &amp; _

12

&quot; SET ServiceRequestNum=&quot; &amp; Me.txtSRnum &amp; _

&quot;, CategoryNum=&#39;&quot; &amp; Me.txtCatNum &amp; &quot;&#39;&quot; &amp; _

&quot;, CustomerNum=&#39;&quot; &amp; Me.txtCnum &amp; &quot;&#39;&quot; &amp; _

&quot;, ServiceDate=&#39;&quot; &amp; Me.txtSdate &amp; &quot;&#39;&quot; &amp; _

&quot;, Description=&#39;&quot; &amp; Me.txtDesc &amp; &quot;&#39;&quot; &amp; _

&quot;, Status=&#39;&quot; &amp; Me.txtStatus &amp; &quot;&#39;&quot; &amp; _

&quot;, EstHours=&#39;&quot; &amp; Me.txtEhours &amp; &quot;&#39;&quot; &amp; _

&quot;, SpentHours=&#39;&quot; &amp; Me.txtShours &amp; &quot;&#39;&quot; &amp; _

&quot;, NextServiceDate=&#39;&quot; &amp; Me.txtNservice &amp; &quot;&#39;&quot; &amp; _

&quot;, ProductNum=&#39;&quot; &amp; Me.txtPnum &amp; &quot;&#39;&quot; &amp; _

&quot;, QuotedPrice=&#39;&quot; &amp; Me.txtQprice &amp; &quot;&#39;&quot; &amp; _

&quot;, NumOrdered=&#39;&quot; &amp; Me.txtNumOr &amp; &quot;&#39;&quot; &amp; _

&quot; WHERE ServiceRequestNum=&quot; &amp; Me.txtSRnum.Tag

End If

&#39;clear form

btnClear_Click()

&#39;refresh data

ServiceRequest_SubForm.Form.Requery()

End Sub

Private Sub btnClear_Click()

Me.txtSRnum = &quot;&quot;

Me.txtCatNum = &quot;&quot;

Me.txtCnum = &quot;&quot;

Me.txtSdate = &quot;&quot;

Me.txtDesc = &quot;&quot;

Me.txtStatus = &quot;&quot;

Me.txtEhours = &quot;&quot;

Me.txtShours = &quot;&quot;

Me.txtNservice = &quot;&quot;

Me.txtPnum = &quot;&quot;

Me.txtQprice = &quot;&quot;

Me.txtNumOr = &quot;&quot;

Me.txtSRnum.SetFocus()

Me.btnEdit.Enabled = True

Me.btnAdd.Caption = &quot;Add&quot;

Me.txtSRnum.Tag = &quot;&quot;

End Sub

Private Sub btnClose_Click()

DoCmd.Close()

End Sub

Private Sub btnDelete_Click()

If Not (Me.ServiceRequest_SubForm.Form.Recordset.EOF And

Me.ServiceRequest_SubForm.Form.Recordset.EOF) Then

If MsgBox(&quot;Are you sure you want to delete?&quot;, vbYesNo) = vbYes Then

CurrentDb.Execute &quot;DELETE FROM ServiceRequest &quot; &amp; _

&quot; WHERE ServiceRequestNum=&quot; &amp;

Me.ServiceRequest_SubForm.Form.Recordset.Fields(&quot;ServiceRequestNum&quot;)

Me.ServiceRequest_SubForm.Form.Requery()

End If

End If

End Sub

13

Private Sub btnEdit_Click()

If Not (Me.ServiceRequest_SubForm.Form.Recordset.EOF And

Me.ServiceRequest_SubForm.Form.Recordset.BOF) Then

With Me.ServiceRequest_SubForm.Form.Recordset

Me.txtSRnum = .Fields(&quot;ServiceRequestNum&quot;)

Me.txtCatNum = .Fields(&quot;CategoryNum&quot;)

Me.txtCnum = .Fields(&quot;CustomerNum&quot;)

Me.txtSdate = .Fields(&quot;ServiceDate&quot;)

Me.txtDesc = .Fields(&quot;Description&quot;)

Me.txtStatus = .Fields(&quot;Status&quot;)

Me.txtEhours = .Fields(&quot;EstHours&quot;)

Me.txtShours = .Fields(&quot;SpentHours&quot;)

Me.txtNservice = .Fields(&quot;NextServiceDate&quot;)

Me.txtPnum = .Fields(&quot;ProductNum&quot;)

Me.txtQprice = .Fields(&quot;QuotedPrice&quot;)

Me.txtNumOr = .Fields(&quot;NumOrdered&quot;)

&#39;store customer number in Tag of txtCnum in case id is modified

Me.txtSRnum.Tag = .Fields(&quot;ServiceRequestNum&quot;)

&#39;change caption of add to update

Me.btnAdd.Caption = &quot;Update&quot;

&#39;disable button edit

Me.btnEdit.Enabled = False

End With

End If

End Sub

Product Form

Option Compare Database

Private Sub btnClear_Click()

Me.txtPnum = &quot;&quot;

Me.txtDesc = &quot;&quot;

Me.txtOhand = &quot;&quot;

Me.txtCat = &quot;&quot;

Me.txtPrice = &quot;&quot;

Me.txtPnum.SetFocus()

Me.btnEdit.Enabled = True

Me.btnAdd.Caption = &quot;Add&quot;

Me.txtPnum.Tag = &quot;&quot;

End Sub

Private Sub btnClose_Click()

DoCmd.Close()

End Sub

Private Sub btnDelete_Click()

If Not (Me.Product_SubForm.Form.Recordset.EOF And

Me.Product_SubForm.Form.Recordset.EOF) Then

If MsgBox(&quot;Are you sure you want to delete?&quot;, vbYesNo) = vbYes Then

CurrentDb.Execute &quot;DELETE FROM Product &quot; &amp; _

&quot; WHERE ProductNum=&quot; &amp;

Me.Product_SubForm.Form.Recordset.Fields(&quot;ProductNum&quot;)

Me.Product_SubForm.Form.Requery()

End If

End If

14

End Sub

Private Sub btnEdit_Click()

If Not (Me.Product_SubForm.Form.Recordset.EOF And

Me.Product_SubForm.Form.Recordset.BOF) Then

With Me.Product_SubForm.Form.Recordset

Me.txtPnum = .Fields(&quot;ProductNum&quot;)

Me.txtDesc = .Fields(&quot;Description&quot;)

Me.txtOhand = .Fields(&quot;OnHand&quot;)

Me.txtCat = .Fields(&quot;Category&quot;)

Me.txtPrice = .Fields(&quot;Price&quot;)

&#39;store customer number in Tag of txtCnum in case id is modified

Me.txtPnum.Tag = .Fields(&quot;ProductNum&quot;)

&#39;change caption of add to update

Me.btnAdd.Caption = &quot;Update&quot;

&#39;disable button edit

Me.btnEdit.Enabled = False

End With

End If

End Sub

Private Sub btnAdd_Click()

&#39;two options when you click on ADD

If Me.txtPnum.Tag &amp; &quot;&quot; = &quot;&quot; Then

&#39;add data

CurrentDb.Execute &quot;INSERT INTO Product(ProductNum, Description,

OnHand, Category, Price) &quot; &amp; _

&quot; VALUES(&quot; &amp; Me.txtPnum &amp; &quot;,&#39;&quot; &amp; Me.txtDesc &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtOhand

&amp; &quot;&#39;,&#39;&quot; &amp; Me.txtCat &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtPrice &amp; &quot;&#39;) &quot;

Else

CurrentDb.Execute &quot;UPDATE Product &quot; &amp; _

&quot; SET ProductNum=&quot; &amp; Me.txtPnum &amp; _

&quot;, Description=&#39;&quot; &amp; Me.txtDesc &amp; &quot;&#39;&quot; &amp; _

&quot;, OnHand=&#39;&quot; &amp; Me.txtOhand &amp; &quot;&#39;&quot; &amp; _

&quot;, Category=&#39;&quot; &amp; Me.txtCat &amp; &quot;&#39;&quot; &amp; _

&quot;, Price=&#39;&quot; &amp; Me.txtPrice &amp; &quot;&#39;&quot; &amp; _

&quot; WHERE ProductNum=&quot; &amp; Me.txtPnum.Tag

End If

&#39;clear form

btnClear_Click()

&#39;refresh data

Product_SubForm.Form.Requery()

End Sub

Service Category Form

Option Compare Database

Private Sub btnAdd_Click()

&#39;two options when you click on ADD

If Me.txtCatNum.Tag &amp; &quot;&quot; = &quot;&quot; Then

&#39;add data

CurrentDb.Execute &quot;INSERT INTO ServiceCategory(CategoryNum,

CategoryDescription) &quot; &amp; _

&quot; VALUES(&quot; &amp; Me.txtCatNum &amp; &quot;,&#39;&quot; &amp; Me.txtCatDesc &amp; &quot;&#39;) &quot;

15

Else

CurrentDb.Execute &quot;UPDATE ServiceCategory &quot; &amp; _

&quot; SET CategoryNum=&quot; &amp; Me.txtCatNum &amp; _

&quot;, CategoryDescription=&#39;&quot; &amp; Me.txtCatDesc &amp; &quot;&#39;&quot; &amp; _

&quot; WHERE CategoryNum=&quot; &amp; Me.txtCatNum.Tag

End If

&#39;clear form

btnClear_Click()

&#39;refresh data

ServiceCategory_SubForm.Form.Requery()

End Sub

Private Sub btnClear_Click()

Me.txtCatNum = &quot;&quot;

Me.txtCatDesc = &quot;&quot;

Me.txtCatNum.SetFocus()

Me.btnEdit.Enabled = True

Me.btnAdd.Caption = &quot;Add&quot;

Me.txtCatNum.Tag = &quot;&quot;

End Sub

Private Sub btnClose_Click()

DoCmd.Close()

End Sub

Private Sub btnDelete_Click()

If Not (Me.ServiceCategory_SubForm.Form.Recordset.EOF And

Me.ServiceCategory_SubForm.Form.Recordset.EOF) Then

If MsgBox(&quot;Are you sure you want to delete?&quot;, vbYesNo) = vbYes Then

CurrentDb.Execute &quot;DELETE FROM ServiceCategory &quot; &amp; _

&quot; WHERE CategoryNum=&quot; &amp;

Me.ServiceCategory_SubForm.Form.Recordset.Fields(&quot;CategoryNum&quot;)

Me.ServiceCategory_SubForm.Form.Requery()

End If

End If

End Sub

Private Sub btnEdit_Click()

If Not (Me.ServiceCategory_SubForm.Form.Recordset.EOF And

Me.ServiceCategory_SubForm.Form.Recordset.BOF) Then

With Me.ServiceCategory_SubForm.Form.Recordset

Me.txtCatNum = .Fields(&quot;CategoryNum&quot;)

Me.txtCatDesc = .Fields(&quot;CategoryDescription&quot;)

&#39;store customer number in Tag of txtCnum in case id is modified

Me.txtCatNum.Tag = .Fields(&quot;CategoryNum&quot;)

&#39;change caption of add to update

Me.btnAdd.Caption = &quot;Update&quot;

&#39;disable button edit

Me.btnEdit.Enabled = False

End With

End If

End Sub

Customer Form

16

Option Compare Database

Private Sub btnAdd_Click()

&#39;two options when you click on ADD

If Me.txtCnum.Tag &amp; &quot;&quot; = &quot;&quot; Then

&#39;add data

CurrentDb.Execute &quot;INSERT INTO Customer(CustomerNum, CustomerName,

Street, City, State, PostalCode, Balance, TechNum) &quot; &amp; _

&quot; VALUES(&quot; &amp; Me.txtCnum &amp; &quot;,&#39;&quot; &amp; Me.txtCname &amp; &quot;&#39;,&#39;&quot; &amp;

Me.txtStreet &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtCity &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtState &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtPostal &amp;

&quot;&#39;,&#39;&quot; &amp; Me.txtBalance &amp; &quot;&#39;,&#39;&quot; &amp; Me.txtTnum &amp; &quot;&#39;) &quot;

Else

CurrentDb.Execute &quot;UPDATE Customer &quot; &amp; _

&quot; SET CustomerNum=&quot; &amp; Me.txtCnum &amp; _

&quot;, CustomerName=&#39;&quot; &amp; Me.txtCname &amp; &quot;&#39;&quot; &amp; _

&quot;, Street=&#39;&quot; &amp; Me.txtStreet &amp; &quot;&#39;&quot; &amp; _

&quot;, City=&#39;&quot; &amp; Me.txtCity &amp; &quot;&#39;&quot; &amp; _

&quot;, State=&#39;&quot; &amp; Me.txtState &amp; &quot;&#39;&quot; &amp; _

&quot;, PostalCode=&#39;&quot; &amp; Me.txtPostal &amp; &quot;&#39;&quot; &amp; _

&quot;, Balance=&#39;&quot; &amp; Me.txtBalance &amp; &quot;&#39;&quot; &amp; _

&quot;, TechNum=&#39;&quot; &amp; Me.txtTnum &amp; &quot;&#39;&quot; &amp; _

&quot; WHERE CustomerNum=&quot; &amp; Me.txtCnum.Tag

End If

&#39;clear form

btnClear_Click()

&#39;refresh data

Customer_Subform.Form.Requery()

End Sub

Private Sub btnClear_Click()

Me.txtCnum = &quot;&quot;

Me.txtCname = &quot;&quot;

Me.txtStreet = &quot;&quot;

Me.txtCity = &quot;&quot;

Me.txtState = &quot;&quot;

Me.txtPostal = &quot;&quot;

Me.txtBalance = &quot;&quot;

Me.txtTnum = &quot;&quot;

Me.txtCnum.SetFocus()

Me.btnEdit.Enabled = True

Me.btnAdd.Caption = &quot;Add&quot;

Me.txtCnum.Tag = &quot;&quot;

End Sub

Private Sub btnClose_Click()

DoCmd.Close()

End Sub

Private Sub btnDelete_Click()

If Not (Me.Customer_Subform.Form.Recordset.EOF And

Me.Customer_Subform.Form.Recordset.EOF) Then

If MsgBox(&quot;Are you sure you want to delete?&quot;, vbYesNo) = vbYes Then

17

CurrentDb.Execute &quot;DELETE FROM Customer &quot; &amp; _

&quot; WHERE CustomerNum=&quot; &amp;

Me.Customer_Subform.Form.Recordset.Fields(&quot;CustomerNum&quot;)

Me.Customer_Subform.Form.Requery()

End If

End If

End Sub

Private Sub btnEdit_Click()

If Not (Me.Customer_Subform.Form.Recordset.EOF And

Me.Customer_Subform.Form.Recordset.BOF) Then

With Me.Customer_Subform.Form.Recordset

Me.txtCnum = .Fields(&quot;CustomerNum&quot;)

Me.txtCname = .Fields(&quot;CustomerName&quot;)

Me.txtStreet = .Fields(&quot;Street&quot;)

Me.txtCity = .Fields(&quot;City&quot;)

Me.txtState = .Fields(&quot;State&quot;)

Me.txtPostal = .Fields(&quot;PostalCode&quot;)

Me.txtBalance = .Fields(&quot;Balance&quot;)

Me.txtTnum = .Fields(&quot;TechNum&quot;)

&#39;store customer number in Tag of txtCnum in case id is modified

Me.txtCnum.Tag = .Fields(&quot;CustomerNum&quot;)

&#39;change caption of add to update

Me.btnAdd.Caption = &quot;Update&quot;

&#39;disable button edit

Me.btnEdit.Enabled = False

End With

End If

End Sub