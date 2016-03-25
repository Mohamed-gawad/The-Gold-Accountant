Imports System.Globalization
Imports System.Drawing.Printing
Imports CrystalDecisions.Shared
Public Class frmPurchases_Bill
#Region "Variable"
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st As String
    Dim x, A, xx, y As Integer
    Dim C As Double ' متغير لحجر قيمة تكلفة الكمية الموجودة من صنف ما
    Dim Best_Cost As Double ' متغير لحجز قيمة السعر المرجح
    Dim par As String ' الباركود
#End Region

#Region "Functions"
    Private Sub New_record()
        Try
            drg.Rows.Clear()
            Myconn.ClearAllControls(GroupBox1, True)
            Myconn.Autonumber("Pur_Bill_num", "Purchases", txtBill_ID, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Filldrg()
        Try
            Select Case x
                Case 0
                    st = " Purchases.Pur_Bill_num = " & CInt(txtBill_ID.Text)
                Case 1
                    st = " Purchases.Pur_Bill_num = " & CInt(txtSearch.Text)
                Case 2
                    st = " "
            End Select
            drg.Rows.Clear()
            Myconn.ExecQuery("SELECT Purchases.Status As Expr1,Purchases.ID, Purchases.Pur_Date, Purchases.Pur_Time, Purchases.Pur_Bill_num, Purchases.Supplier_ID, Purchases.items_Cod, (Purchases.Customer_Price) as factory_Price, Purchases.Reduce, Purchases.Pur_Price, Purchases.Iteme_Number, Purchases.Total_Price, Purchases.Stock_ID, Purchases.Status,(Items.Total_Price) as Gomla,Items.Items_Name,Items.Parcode,Items.Customer_Price, Employees.Employee_Name
                            FROM (Employees RIGHT JOIN Users_ID ON Employees.Employee_ID = Users_ID.Employee_ID) RIGHT JOIN (Purchases LEFT JOIN Items ON Purchases.items_Cod = Items.items_Cod) ON Users_ID.Employee_ID = Purchases.Employee_ID
                            where " & st & " order by Purchases.ID")

            If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
            Dim V1 As Double = 0
            Dim V2 As Double = 0
            Dim B As Double = 0
            For i As Integer = 0 To Myconn.dt.Rows.Count - 1
                Dim r As DataRow = Myconn.dt.Rows(i)
                drg.Rows.Add()
                drg.Rows(i).Cells(0).Value = i + 1
                drg.Rows(i).Cells(1).Value = r("Pur_Time") ' الوفت
                drg.Rows(i).Cells(2).Value = r("Items_Name") ' الصنف
                drg.Rows(i).Cells(3).Value = r("Parcode") ' الباركود
                drg.Rows(i).Cells(4).Value = If(My.Settings.Factory_Price = True, r("factory_Price"), r("Customer_Price")) 'سعر المصنع أو المستهلك
                drg.Rows(i).Cells(5).Value = If(My.Settings.Factory_Price = True, r("Reduce") & " % ", r("Gomla")) ' سعر الجملة او الخصم
                drg.Rows(i).Cells(6).Value = r("Pur_Price") ' سعر التكلفة
                drg.Rows(i).Cells(7).Value = r("Iteme_Number") ' العدد
                drg.Rows(i).Cells(8).Value = r("Total_Price") ' اجمالي التكلفة
                drg.Rows(i).Cells(9).Value = If(My.Settings.Factory_Price = True, r("factory_Price") * r("Iteme_Number"), r("Customer_Price") * r("Iteme_Number")) ' إجمالي المصنع أو المستهلك
                drg.Rows(i).Cells(10).Value = r("Customer_Price") ' سعر المستهلك
                drg.Rows(i).Cells(11).Value = r("Status") ' الحالة
                drg.Rows(i).Cells(12).Value = r("ID") ' رقم السجل
                drg.Rows(i).Cells(13).Value = r("items_Cod") ' كود الصنف
                drg.Rows(i).Cells(14).Value = r("Gomla") ' سعر الجملة
                drg.Rows(i).Cells(15).Value = r("Employee_Name") ' المستخدم

                If drg.Rows(i).Cells(11).Value = True Then
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.LemonChiffon
                    V1 += CDec(drg.Rows(i).Cells(9).Value)
                    V2 += CDec(drg.Rows(i).Cells(8).Value)
                Else
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    B += CDec(drg.Rows(i).Cells(8).Value)
                End If
            Next

            Myconn.DataGridview_MoveLast(drg, 2)
            Label19.Text = Math.Round(V2, 2) ' التكلفة
            Label18.Text = Math.Round(V1, 2) ' المصنع
            Label22.Text = Math.Round(B, 2) '  المرتجع
            Label20.Text = Math.Round(((Val(Val(Label18.Text) - Val(Label19.Text)) / Val(Label18.Text)) * 100), 2) & " % "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Binding()
        Try
            Myconn.ExecQuery("SELECT Purchases.items_Cod, (Purchases.Customer_Price) as factory_Price, Purchases.Reduce, Purchases.Pur_Price, Purchases.Iteme_Number, Purchases.Total_Price, (Items.Customer_Price) as sales_Price,(items.Total_Price) as Gomla_Price , Items.Parcode
                                 FROM Items RIGHT JOIN Purchases ON Items.items_Cod = Purchases.items_Cod
                                 where Purchases.ID =" & CInt(drg.CurrentRow.Cells(12).Value) & "")

            Dim r As DataRow = Myconn.dt.Rows(0)
            Dim F As Double = If(IsDBNull(r("factory_Price")), 0, r("factory_Price"))
            Dim S As Double = If(IsDBNull(r("sales_Price")), 0, r("sales_Price"))
            Dim G As Double = If(IsDBNull(r("Gomla_Price")), 0, r("Gomla_Price"))

            cbo_Kind.SelectedValue = r("items_Cod")
            txt_Barcode.Text = r("Parcode")
            txt_Factory.Text = r("factory_Price")
            txt_Reduce.Text = r("Reduce")
            txt_Cost.Text = r("Pur_Price")
            txt_Amount.Text = r("Iteme_Number")
            txt_Total.Text = r("Total_Price")
            txt_Gomla.Text = If(My.Settings.Customer_Price = True, If(G <= 0, 0, Math.Round((((G - F) / F) * 100), 2)), G)
            txt_Customer2.Text = If(My.Settings.Customer_Price = True, If(S <= 0, 0, Math.Round((((S - F) / F) * 100), 2)), S)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Save_recod()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@Pur_Date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
                .Addparam("@Pur_Time", Label14.Text)
                .Addparam("@Pur_Bill_num", txtBill_ID.Text)
                .Addparam("@Supplier_ID", cbo_Supplier.SelectedValue)
                .Addparam("@items_Cod", cbo_Kind.SelectedValue)
                .Addparam("@Customer_Price", txt_Factory.Text)
                .Addparam("@Reduce", txt_Reduce.Text)
                .Addparam("@Pur_Price", txt_Cost.Text)
                .Addparam("@Iteme_Number", txt_Amount.Text)
                .Addparam("@Total_Price", txt_Total.Text)
                .Addparam("@Stock_ID", cbo_Stock.SelectedValue)
                .Addparam("@Status", 1)
                .Addparam("@Employee_ID", My.Settings.user_ID)

                .ExecQuery("insert into  [Purchases] (Pur_Date, Pur_Time, Pur_Bill_num, Supplier_ID, items_Cod, Customer_Price, Reduce, Pur_Price, Iteme_Number, Total_Price, Stock_ID, Status, Employee_ID)
                                           values(@Pur_Date,@Pur_Time,@Pur_Bill_num,@Supplier_ID,@items_Cod,@Customer_Price,@Reduce,@Pur_Price,@Iteme_Number,@Total_Price,@Stock_ID,@Status,@Employee_ID)")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Update_record()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@Pur_Date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
                .Addparam("@Pur_Time", Label14.Text)
                .Addparam("@Pur_Bill_num", txtBill_ID.Text)
                .Addparam("@Supplier_ID", cbo_Supplier.SelectedValue)
                .Addparam("@items_Cod", cbo_Kind.SelectedValue)
                .Addparam("@Customer_Price", txt_Factory.Text)
                .Addparam("@Reduce", txt_Reduce.Text)
                .Addparam("@Pur_Price", txt_Cost.Text)
                .Addparam("@Iteme_Number", txt_Amount.Text)
                .Addparam("@Total_Price", txt_Total.Text)
                .Addparam("@Stock_ID", cbo_Stock.SelectedValue)
                .Addparam("@Employee_ID", 1)
                .Addparam("@ID", drg.CurrentRow.Cells(12).Value)
                .ExecQuery("Update [Purchases] set Pur_Date=@Pur_Date, Pur_Time=@Pur_Time, Pur_Bill_num=@, Supplier_ID=@Supplier_ID, items_Cod=@items_Cod, 
                                               Customer_Price=@Customer_Price, Reduce=@Reduce, Pur_Price=@Pur_Price, Iteme_Number=@Iteme_Number,
                                               Total_Price=@Total_Price, Stock_ID=@Stock_ID, Employee_ID=@Employee_ID where ID = @ID")
                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Update_Items_record()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@Customer_Price", If(My.Settings.Customer_Price = False, txt_Customer2.Text, Val(txt_Factory.Text) * (1 + Val(txt_Customer2.Text / 100)))) ' سعر المستهلك
                .Addparam("@Total_Price", If(My.Settings.Customer_Price = False, txt_Gomla.Text, Val(txt_Factory.Text) * (1 + Val(txt_Gomla.Text / 100)))) ' سعر الجملة
                .Addparam("@cost_Price", Best_Cost) ' التكلفة
                .Addparam("@Supplier_ID", cbo_Supplier.SelectedValue) ' المورد
                .Addparam("@items_Cod", cbo_Kind.SelectedValue) ' الصنف
                .ExecQuery("Update [Items] set Customer_Price=@Customer_Price,Total_Price=@Total_Price,cost_Price=@cost_Price,Supplier_ID=@Supplier_ID  where items_Cod = @items_Cod")
                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
#End Region

    Private Sub frmPurchases_Bill_Activated(sender As Object, e As EventArgs) Handles Me.Activated
#Region "Customer_Price"
        If My.Settings.Customer_Price = True Then
            Label11.Text = "نسبة البيع للجملة"
            Label12.Text = "نسبة البيع للمستهلك"
        Else
            Label11.Text = "سعر البيع للجملة"
            Label12.Text = "سعر البيع للمستهلك"
        End If
#End Region

        cbo_Kind.DataSource = Nothing
        fin = False
        Myconn.Fillcombo("select Items_Name,items_Cod from [items] order by [Items_Name]", "[items]", "items_Cod", "Items_Name", Me, cbo_Kind)
        fin = True
    End Sub

    Private Sub frmPurchases_Bill_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width
#Region "Factory_Price"
            If My.Settings.Factory_Price = True Then
                txt_Reduce.Visible = True
                txt_Factory.Visible = True
                txt_Cost.Enabled = False
                Label6.Visible = True
                Label7.Visible = True
                Label15.Text = "سعر المصنع     ="
                RB1.Visible = True
                RB2.Visible = True
                RB1.Checked = True
                drg.Columns(4).HeaderText = "سعر المصنع"
                drg.Columns(5).HeaderText = "نسبة الخصم"
                drg.Columns(9).HeaderText = " إجمالي المصنع"
                drg.Columns(10).Visible = True
                drg.Columns(14).Visible = True
                Label17.Visible = True
                Label20.Visible = True

                txt_Reduce.Top = 74
                txt_Factory.Top = 74
                Label6.Top = 75
                Label7.Top = 75

                txt_Amount.Top = 103
                txt_Cost.Top = 103
                Label9.Top = 105
                Label8.Top = 105

                txt_Total.Top = 132
                txtBest_Price.Top = 132
                Label23.Top = 135
                Label10.Top = 135

                txt_Customer2.Top = 161
                txt_Gomla.Top = 161
                Label11.Top = 165
                Label12.Top = 165

            Else
                txt_Reduce.Visible = False
                txt_Factory.Visible = False
                txt_Cost.Enabled = True
                Label6.Visible = False
                Label7.Visible = False
                Label15.Text = "سعر المستهلك    ="
                RB1.Visible = False
                RB2.Visible = False
                RB2.Checked = True
                drg.Columns(4).HeaderText = "سعر المستهلك"
                drg.Columns(5).HeaderText = "سعر الجملة"
                drg.Columns(9).HeaderText = "إجمالي المستهلك"
                drg.Columns(10).Visible = False
                drg.Columns(14).Visible = False
                Label17.Visible = False
                Label20.Visible = False

                txt_Amount.Top = 74
                txt_Cost.Top = 74
                Label9.Top = 75
                Label8.Top = 75

                txt_Total.Top = 102
                txtBest_Price.Top = 102
                Label23.Top = 105
                Label10.Top = 105

                txt_Customer2.Top = 132
                txt_Gomla.Top = 132
                Label11.Top = 135
                Label12.Top = 135


            End If
#End Region
#Region "Customer_Price"
            If My.Settings.Customer_Price = True Then
                Label11.Text = "نسبة البيع للجملة"
                Label12.Text = "نسبة البيع للمستهلك"
            Else
                Label11.Text = "سعر البيع للجملة"
                Label12.Text = "سعر البيع للمستهلك"
            End If
#End Region

            If F <> 1 Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Return
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnBack.Enabled = r("U_back").ToString
                    btnSave.Enabled = r("U_add").ToString
                    btnSearch.Enabled = r("U_search").ToString
                    btnUpdat.Enabled = r("U_updat").ToString
                    btnNew.Enabled = r("U_new").ToString
                    btnDel.Enabled = r("U_delete").ToString
                    btnPrint.Enabled = r("U_print").ToString
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        fin = False
        Myconn.Fillcombo("select Items_Name,items_Cod from [items] order by [Items_Name]", "[items]", "items_Cod", "Items_Name", Me, cbo_Kind)
        Myconn.Fillcombo("select * from Supplier order by Supplier_Name", "Supplier", "Supplier_ID", "Supplier_Name", Me, cbo_Supplier)
        Myconn.Fillcombo("select * from [Stocks] order by Stock_ID", "[Stocks]", "Stock_ID", "Stock_Name", Me, cbo_Stock)
        fin = True
        New_record()
        Timer1.Start()

        '-------------------------------------------------------------------------------------------------- النسخة التجريبية
        'Myconn.ExecQuery("select * from Purchases")
        'If Myconn.Recodcount > 300 Then
        '    MsgBox("هذه النسخة تجريبية")
        '    btnSave.Enabled = False
        '    btnNew.Enabled = False
        '    Return
        'End If
    End Sub
#Region "Buttons"
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        New_record()
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            For Each txt As Control In GroupBox1.Controls
                If TypeOf txt Is TextBox Then
                    If txt.Text = "" Then
                        ErrorProvider1.SetError(txt, "أكمل البيانات")
                        MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                        Return
                    End If
                ElseIf TypeOf txt Is ComboBox Then
                    If txt.Text = "" Then
                        ErrorProvider1.SetError(txt, "أكمل البيانات")
                        MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                        Return
                    End If
                End If
            Next
            For Each txt As Control In GroupBox2.Controls
                If TypeOf txt Is TextBox Then
                    If txt.Text = "" Then
                        ErrorProvider1.SetError(txt, "أكمل البيانات")
                        MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                        Return
                    End If
                ElseIf TypeOf txt Is ComboBox Then
                    If txt.Text = "" Then
                        ErrorProvider1.SetError(txt, "أكمل البيانات")
                        MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                        Return
                    End If
                End If
            Next

            If My.Settings.Customer_Price = False Then
                If Val(txt_Gomla.Text) < Val(txtBest_Price.Text) Then
                    ErrorProvider1.SetError(txt_Gomla, "أكمل البيانات")
                    MsgBox("سعر بيع الجملة غير صحيح", MsgBoxStyle.Critical, "رسالة")
                    Return
                End If
                If Val(txt_Customer2.Text) < Val(txtBest_Price.Text) Then
                    ErrorProvider1.SetError(txt_Customer2, "أكمل البيانات")
                    MsgBox("سعر بيع المستهلك غير صحيح", MsgBoxStyle.Critical, "رسالة")
                    Return
                End If
            Else
                If Val(txtBest_Price.Text) * (1 + Val(txt_Gomla.Text / 100)) < Val(txtBest_Price.Text) Then
                    ErrorProvider1.SetError(txt_Gomla, "أكمل البيانات")
                    MsgBox("سعر بيع الجملة غير صحيح", MsgBoxStyle.Critical, "رسالة")
                    Return
                End If
                If Val(txtBest_Price.Text) * (1 + Val(txt_Customer2.Text / 100)) < Val(txtBest_Price.Text) Then
                    ErrorProvider1.SetError(txt_Customer2, "أكمل البيانات")
                    MsgBox("سعر بيع المستهلك غير صحيح", MsgBoxStyle.Critical, "رسالة")
                    Return
                End If
            End If

            Save_recod()
            Update_Items_record()
            x = 0
            Filldrg()
            cbo_Kind_SelectedIndexChanged(Nothing, Nothing)
            MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
            Myconn.ClearAllControls(GroupBox2, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub btnUpdat_Click(sender As Object, e As EventArgs) Handles btnUpdat.Click
        Try
            For Each txt As Control In GroupBox1.Controls
                If TypeOf txt Is TextBox Then
                    If txt.Text = "" Then
                        ErrorProvider1.SetError(txt, "أكمل البيانات")
                        MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                        Return
                    End If
                ElseIf TypeOf txt Is ComboBox Then
                    If txt.Text = "" Then
                        ErrorProvider1.SetError(txt, "أكمل البيانات")
                        MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                        Return
                    End If
                End If
            Next
            For Each txt As Control In GroupBox2.Controls
                If TypeOf txt Is TextBox Then
                    If txt.Text = "" Then
                        ErrorProvider1.SetError(txt, "أكمل البيانات")
                        MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                        Return
                    End If
                ElseIf TypeOf txt Is ComboBox Then
                    If txt.Text = "" Then
                        ErrorProvider1.SetError(txt, "أكمل البيانات")
                        MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                        Return
                    End If
                End If
            Next

            If My.Settings.Customer_Price = False Then
                If Val(txt_Gomla.Text) < Val(txtBest_Price.Text) Then
                    ErrorProvider1.SetError(txt_Gomla, "أكمل البيانات")
                    MsgBox("سعر بيع الجملة غير صحيح", MsgBoxStyle.Critical, "رسالة")
                    Return
                End If
                If Val(txt_Customer2.Text) < Val(txtBest_Price.Text) Then
                    ErrorProvider1.SetError(txt_Customer2, "أكمل البيانات")
                    MsgBox("سعر بيع المستهلك غير صحيح", MsgBoxStyle.Critical, "رسالة")
                    Return
                End If
            Else
                If Val(txtBest_Price.Text) * (1 + Val(txt_Gomla.Text / 100)) < Val(txtBest_Price.Text) Then
                    ErrorProvider1.SetError(txt_Gomla, "أكمل البيانات")
                    MsgBox("سعر بيع الجملة غير صحيح", MsgBoxStyle.Critical, "رسالة")
                    Return
                End If
                If Val(txtBest_Price.Text) * (1 + Val(txt_Customer2.Text / 100)) < Val(txtBest_Price.Text) Then
                    ErrorProvider1.SetError(txt_Customer2, "أكمل البيانات")
                    MsgBox("سعر بيع المستهلك غير صحيح", MsgBoxStyle.Critical, "رسالة")
                    Return
                End If
            End If

            Update_record()
            Update_Items_record()
            '------------------------------------------------------------------------------------------------------------------------
            Myconn.ExecQuery("SELECT Purchases.Status As Expr1,Purchases.ID, Purchases.Pur_Date, Purchases.Pur_Time, Purchases.Pur_Bill_num, Purchases.Supplier_ID, Purchases.items_Cod, (Purchases.Customer_Price) as factory_Price, Purchases.Reduce, Purchases.Pur_Price, Purchases.Iteme_Number, Purchases.Total_Price, Purchases.Stock_ID, Purchases.Status,(Items.Total_Price) as Gomla,Items.Items_Name,Items.Parcode,Items.Customer_Price, Employees.Employee_Name
                            FROM (Employees RIGHT JOIN Users_ID ON Employees.Employee_ID = Users_ID.Employee_ID) RIGHT JOIN (Purchases LEFT JOIN Items ON Purchases.items_Cod = Items.items_Cod) ON Users_ID.Employee_ID = Purchases.Employee_ID
                            where Purchases.ID =" & CInt(drg.CurrentRow.Cells(12).Value))

            If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub

            Dim r As DataRow = Myconn.dt.Rows(0)
            drg.CurrentRow.Cells(1).Value = r("Pur_Time") ' الوفت
            drg.CurrentRow.Cells(2).Value = r("Items_Name") ' الصنف
            drg.CurrentRow.Cells(3).Value = r("Parcode") ' الباركود
            drg.CurrentRow.Cells(4).Value = If(My.Settings.Factory_Price = True, r("factory_Price"), r("Customer_Price")) 'سعر المصنع أو المستهلك
            drg.CurrentRow.Cells(5).Value = If(My.Settings.Factory_Price = True, r("Reduce") & " % ", r("Gomla")) ' سعر الجملة او الخصم
            drg.CurrentRow.Cells(6).Value = r("Pur_Price") ' سعر التكلفة
            drg.CurrentRow.Cells(7).Value = r("Iteme_Number") ' العدد
            drg.CurrentRow.Cells(8).Value = r("Total_Price") ' اجمالي التكلفة
            drg.CurrentRow.Cells(9).Value = If(My.Settings.Factory_Price = True, r("factory_Price") * r("Iteme_Number"), r("Customer_Price") * r("Iteme_Number")) ' إجمالي المصنع أو المستهلك
            drg.CurrentRow.Cells(10).Value = r("Customer_Price") ' سعر المستهلك
            drg.CurrentRow.Cells(11).Value = r("Status") ' الحالة
            drg.CurrentRow.Cells(12).Value = r("ID") ' رقم السجل
            drg.CurrentRow.Cells(13).Value = r("items_Cod") ' كود الصنف
            drg.CurrentRow.Cells(14).Value = r("Gomla") ' سعر الجملة
            drg.CurrentRow.Cells(15).Value = r("Employee_Name") ' المستخدم

            Myconn.Sum_drg2(drg, 9, Label18)
            Myconn.Sum_drg2(drg, 8, Label19)
            Label20.Text = Math.Round(((Val(Val(Label18.Text) - Val(Label19.Text)) / Val(Label18.Text)) * 100), 2) & " % "
            cbo_Kind_SelectedIndexChanged(Nothing, Nothing)
            MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
            'Account_Cost()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        Try
            If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
            Account_Cost()
            If Val(A - drg.CurrentRow.Cells(7).Value) = 0 Then
                GoTo SS
            Else
                Best_Cost = Math.Round((Val(C - drg.CurrentRow.Cells(8).Value) / Val(A - drg.CurrentRow.Cells(7).Value)), 2)
                Update_Items_record()
            End If

SS:
            With Myconn
                .Addparam("@ID", drg.CurrentRow.Cells(12).Value)
                .ExecQuery("delete from [Purchases] where ID = @ID ")
            End With
            If Myconn.NoErrors(True) = False Then Exit Sub
            drg.Rows.Remove(drg.SelectedRows(0))
            cbo_Kind_SelectedIndexChanged(Nothing, Nothing)
            Myconn.ClearAllControls(GroupBox2, True)
            Myconn.DataGridview_MoveLast(drg, 2)
            Myconn.Sum_drg2(drg, 9, Label18)
            Myconn.Sum_drg2(drg, 8, Label19)
            Label20.Text = Math.Round(((Val(Val(Label18.Text) - Val(Label19.Text)) / Val(Label18.Text)) * 100), 2)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Public Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try
            Myconn.ExecQuery("select Pur_Date,Pur_Bill_num,Supplier_ID,Stock_ID from Purchases where Pur_Bill_num =" & CInt(txtSearch.Text) & "")

            If Myconn.dt.Rows.Count = 0 Then
                MsgBox("لا توجد فاتورة بهذا الرقم ", MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة")
                Return
            End If
            Dim r As DataRow = Myconn.dt.Rows(0)
            D_date.Text = r("Pur_Date").ToString
            txtBill_ID.Text = r("Pur_Bill_num").ToString
            cbo_Supplier.SelectedValue = r("Supplier_ID")
            cbo_Stock.SelectedValue = r("Stock_ID")
            x = 1
            Filldrg()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Try


            If drg.CurrentRow.Cells(11).Value = True Then
                With Myconn
                    .Parames.Clear()
                    .Addparam("@Status", False)
                    .Addparam("@ID", drg.CurrentRow.Cells(12).Value)
                End With
                Myconn.ExecQuery(" Update  Purchases set Status = @Status where ID = @ID")
                drg.CurrentRow.DefaultCellStyle.BackColor = Color.Red
                drg.CurrentRow.Cells(11).Value = False
                Account_Cost()
                Best_Cost = Math.Round((Val(C - drg.CurrentRow.Cells(8).Value) / Val(A - drg.CurrentRow.Cells(7).Value)), 2)
                Update_Items_record()
            Else
                With Myconn
                    .Parames.Clear()
                    .Addparam("@Status", True)
                    .Addparam("@ID", drg.CurrentRow.Cells(12).Value)
                End With
                Myconn.ExecQuery(" Update  Purchases set Status = @Status where ID = @ID")
                drg.CurrentRow.DefaultCellStyle.BackColor = Color.LemonChiffon
                drg.CurrentRow.Cells(11).Value = True
                Account_Cost()
                Best_Cost = Math.Round((Val(C + drg.CurrentRow.Cells(8).Value) / Val(A + drg.CurrentRow.Cells(7).Value)), 2)
                Update_Items_record()
            End If


            Dim V1 As Double = 0
            Dim V2 As Double = 0
            Dim B As Double = 0
            For i As Integer = 0 To drg.Rows.Count - 1
                If drg.Rows(i).Cells(11).Value = True Then
                    V1 += CDec(drg.Rows(i).Cells(9).Value)
                    V2 += CDec(drg.Rows(i).Cells(8).Value)
                Else
                    B += CDec(drg.Rows(i).Cells(8).Value)
                End If
            Next i
            Label22.Text = B
            Label19.Text = V2
            Label18.Text = V1
            Label20.Text = Math.Round(((Val(Val(Label18.Text) - Val(Label19.Text)) / Val(Label18.Text)) * 100), 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub btnAdd_Kind_Click(sender As Object, e As EventArgs) Handles btnAdd_Kind.Click
        Try
            frmAdd_Kind.MdiParent = FrmMain
            frmAdd_Kind.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Select Case cboPrint.SelectedIndex
            Case 0 '  الفاتورة
                If My.Settings.Factory_Price = True Then
                    Print_Pur_Bill()
                Else
                    Print_Pur_Bill_2()
                End If

            Case 1 ' مرتجع الفاتورة
                If My.Settings.Factory_Price = True Then
                    Print_bill_back()
                Else
                    Print_bill_back_2()
                End If

            Case 2 ' باركود الصنف
                Print_One_Barcode(drg)
            Case 3 ' باركود الفاتورة
                Print_Barcode(drg)
            Case 4 ' باركود عدد معين
                Print_Number_Barcode(drg)
        End Select
    End Sub
#End Region

#Region "TextBox"
    Private Sub txt_Barcode_Enter(sender As Object, e As EventArgs) Handles txt_Barcode.Enter
        txt_Barcode.Text = ""
    End Sub
    Private Sub txt_Customer1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_Factory.KeyPress, txt_Amount.KeyPress, txt_Cost.KeyPress, txt_Factory.KeyPress, txt_Customer2.KeyPress, txt_Gomla.KeyPress, txt_Reduce.KeyPress, txt_Total.KeyPress, txtBest_Price.KeyPress
        Myconn.NumberOnly(txt_Amount, e)
    End Sub
    Private Sub txt_Reduce_TextChanged(sender As Object, e As EventArgs) Handles txt_Reduce.TextChanged ' الخصم
        ErrorProvider1.Clear()
        Try
            If RB1.Checked = True Then
                txt_Cost.Text = Math.Round((Val(txt_Factory.Text) - Math.Round((Val(txt_Factory.Text) * Val(txt_Reduce.Text) / 100), 2)), 2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub txt_Cost_TextChanged(sender As Object, e As EventArgs) Handles txt_Cost.TextChanged, txtBest_Price.TextChanged ' سعر التكلفة
        Try
            If My.Settings.Factory_Price = True Then
                If RB2.Checked = True Then
                    txt_Reduce.Text = Math.Round((Val(Val(txt_Factory.Text) - Val(txt_Cost.Text)) * 100 / Val(txt_Factory.Text)), 2)
                End If
            Else
                txt_Factory.Text = txt_Cost.Text
                txt_Reduce.Text = 0
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub txt_Amount_TextChanged(sender As Object, e As EventArgs) Handles txt_Amount.TextChanged ' التكلفة للكمية
        ErrorProvider1.Clear()
        Try
            If txt_Amount.Text = "" Then Return
            txt_Total.Text = Math.Round((Val(txt_Cost.Text) * Val(txt_Amount.Text)), 2) ' إجمالي التكلفة
            txtBest_Price.Text = If(Val(A + txt_Amount.Text) > 0, Math.Round((Val((C + txt_Total.Text)) / Val((A + txt_Amount.Text))), 2), 0) ' السعر المرجح
            Best_Cost = txtBest_Price.Text ' متغير لحجز قيمة السعر المرجح
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub txt_Factory_Enter(sender As Object, e As EventArgs) Handles txt_Factory.Enter
        txt_Factory.Text = Nothing
        txt_Amount.Text = Nothing
        txt_Reduce.Text = Nothing
        txt_Total.Text = Nothing
        txtBest_Price.Text = Nothing
        txt_Gomla.Text = Nothing
        txt_Customer2.Text = Nothing
    End Sub
    Private Sub txt_Reduce_Enter(sender As Object, e As EventArgs) Handles txt_Reduce.Enter
        txt_Reduce.Text = Nothing
        txt_Cost.Text = Nothing
        txt_Amount.Text = Nothing
        txt_Total.Text = Nothing
        txtBest_Price.Text = Nothing
        txt_Gomla.Text = Nothing
        txt_Customer2.Text = Nothing
    End Sub

    Private Sub txt_Cost_Enter(sender As Object, e As EventArgs) Handles txt_Cost.Enter
        txt_Reduce.Text = Nothing
        txt_Amount.Text = Nothing
        txt_Cost.Text = Nothing
        txt_Gomla.Text = Nothing
        txt_Customer2.Text = Nothing
        txt_Total.Text = Nothing
        txtBest_Price.Text = Nothing
    End Sub
    Private Sub txt_Amount_Enter(sender As Object, e As EventArgs) Handles txt_Amount.Enter
        txt_Amount.Text = Nothing
        txt_Gomla.Text = Nothing
        txt_Customer2.Text = Nothing
        txt_Total.Text = Nothing
        txtBest_Price.Text = Nothing
    End Sub
    Private Sub txt_Gomla_Enter(sender As Object, e As EventArgs) Handles txt_Gomla.Enter
        txt_Gomla.Text = Nothing
    End Sub
    Private Sub txt_Gomla_TextChanged(sender As Object, e As EventArgs) Handles txt_Gomla.TextChanged
        ErrorProvider1.Clear()
        'If (My.Settings.Customer_Price = False, txt_Gomla.Text, Val(txt_Factory.Text) * (1 + Val(txt_Gomla.Text / 100)))
    End Sub
    Private Sub txt_Customer2_Enter(sender As Object, e As EventArgs) Handles txt_Customer2.Enter
        txt_Customer2.Text = Nothing
    End Sub
    Private Sub txtSearch_KeyUp(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyUp
        If e.KeyCode = Keys.Enter Then
            btnSearch_Click(Nothing, Nothing)
        End If
    End Sub
    Private Sub txt_Customer2_TextChanged(sender As Object, e As EventArgs) Handles txt_Customer2.TextChanged
        ErrorProvider1.Clear()
    End Sub


#Region "Moving"
    Private Sub txt_Barcode_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_Barcode.KeyUp ' الباركود
        Try
            If txt_Barcode.Text = "" Then Return
            If cbo_Stock.SelectedIndex = -1 Then MsgBox("قم باختيار المخزن الذي سيتم اضافة البضاعة اليه", MsgBoxStyle.Critical, "رسالة") : Return
            If e.KeyCode = Keys.Enter = True Then
                Myconn.ExecQuery("Select items_Cod from items where Parcode Like '" & txt_Barcode.Text & "'")
                If Myconn.dt.Rows.Count = 0 Then
                    cbo_Kind.SelectedIndex = -1
                    MsgBox("الصنف غير موجود أو الباركود غير صحيح", MsgBoxStyle.Critical, "رسالة")

                    Return
                End If
                Dim r As DataRow = Myconn.dt.Rows(0)
                cbo_Kind.SelectedValue = If(IsDBNull(r("items_Cod")), -1, r("items_Cod"))

                If My.Settings.Factory_Price = True Then
                    txt_Factory.Focus()
                Else
                    txt_Cost.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
            Return
        End Try

    End Sub
    Private Sub cbo_Kind_KeyUp(sender As Object, e As KeyEventArgs) Handles cbo_Kind.KeyUp
        If e.KeyCode = Keys.Enter = True Then
            If My.Settings.Factory_Price = True Then
                txt_Factory.Focus()
            Else
                txt_Cost.Focus()
            End If
        End If
    End Sub
    Private Sub txt_Factory_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_Factory.KeyUp
        If e.KeyCode = Keys.Enter = True Then
            If RB1.Checked = True Then
                txt_Reduce.Focus()
            Else
                txt_Cost.Focus()
            End If

        End If
    End Sub
    Private Sub txt_Reduce_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_Reduce.KeyUp
        If e.KeyCode = Keys.Enter = True Then
            txt_Amount.Focus()
        End If
    End Sub
    Private Sub txt_Cost_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_Cost.KeyUp
        If e.KeyCode = Keys.Enter = True Then
            txt_Amount.Focus()
        End If
    End Sub
    Private Sub txt_Amount_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_Amount.KeyUp
        If e.KeyCode = Keys.Enter = True Then
            txt_Gomla.Focus()
        End If
    End Sub
    Private Sub txt_Gomla_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_Gomla.KeyUp
        If e.KeyCode = Keys.Enter = True Then
            txt_Customer2.Focus()
        End If
    End Sub
    Private Sub txt_Customer2_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_Customer2.KeyUp
        If e.KeyCode = Keys.Enter Then
            If btnSave.Enabled = False Then MsgBox("هذه النسخة تجريبية", MsgBoxStyle.Critical, "رسالة") : Return
            btnSave_Click(Nothing, Nothing)
            txt_Barcode.Focus()
        End If
    End Sub

#End Region
#End Region
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label14.Text = TimeOfDay.ToString("hh:mm:ss tt", CultureInfo.CreateSpecificCulture("ar-eg"))
    End Sub
    Private Sub RB1_CheckedChanged(sender As Object, e As EventArgs) Handles RB1.CheckedChanged
        If RB1.Checked = True Then
            txt_Reduce.Enabled = True
            txt_Cost.Enabled = False
        Else
            txt_Reduce.Enabled = False
            txt_Cost.Enabled = True
        End If
    End Sub
    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Binding()
        Account_Cost()
    End Sub
    Private Sub drg_MouseClick(sender As Object, e As MouseEventArgs) Handles drg.MouseClick
        If (e.Button = MouseButtons.Right) Then
            ContextMenuStrip1.Show(drg, e.Location)
        End If
    End Sub
    Private Sub drg_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellDoubleClick
        drg.CurrentRow.Selected = False
    End Sub
    Private Sub Back_amount_Click(sender As Object, e As EventArgs) Handles Back_amount.Click
        Try
            xx = InputBox("أدخل الكمية المرتجعة", "إرجاع كمية محددة")
        Catch ex As Exception
            Return
        End Try

        If x >= drg.CurrentRow.Cells(7).Value Then
            MsgBox("الكمية المرتجعة أكبر من الكمية الموجودة")
            Return
        End If
        With Myconn
            .Parames.Clear()
            .Addparam("@Iteme_Number", drg.CurrentRow.Cells(7).Value - xx)
            .Addparam("@ID", drg.CurrentRow.Cells(12).Value)
        End With
        Myconn.ExecQuery(" Update  Purchases set Iteme_Number = @Iteme_Number where ID = @ID")
        '--------------------------------------------------------------------------------------------------------------- اضافة الكمية المرتجعة
        With Myconn
            .Parames.Clear()
            .Addparam("@Pur_Date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
            .Addparam("@Pur_Time", Label14.Text)
            .Addparam("@Pur_Bill_num", txtBill_ID.Text)
            .Addparam("@Supplier_ID", cbo_Supplier.SelectedValue)
            .Addparam("@items_Cod", cbo_Kind.SelectedValue)
            .Addparam("@Customer_Price", txt_Factory.Text)
            .Addparam("@Reduce", txt_Reduce.Text)
            .Addparam("@Pur_Price", txt_Cost.Text)
            .Addparam("@Iteme_Number", xx)
            .Addparam("@Total_Price", Val(txt_Cost.Text) * xx)
            .Addparam("@Stock_ID", cbo_Stock.SelectedValue)
            .Addparam("@Status", 0)
            .Addparam("@Employee_ID", 1)

            .ExecQuery("insert into  [Purchases] (Pur_Date, Pur_Time, Pur_Bill_num, Supplier_ID, items_Cod, Customer_Price, Reduce, Pur_Price, Iteme_Number, Total_Price, Stock_ID, Status, Employee_ID)
                                           values(@Pur_Date,@Pur_Time,@Pur_Bill_num,@Supplier_ID,@items_Cod,@Customer_Price,@Reduce,@Pur_Price,@Iteme_Number,@Total_Price,@Stock_ID,@Status,@Employee_ID)")

            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
        x = 0
        Filldrg()
    End Sub
    Private Sub cbo_Kind_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_Kind.SelectedIndexChanged
        Try
            If Not fin Then Return
            Account_Cost()
            txt_Barcode.Text = par
            Myconn.ExecQuery("Select * from Items where items_Cod =" & CInt(cbo_Kind.SelectedValue))
            Dim r As DataRow = Myconn.dt.Rows(0)
            txtKind.Text = r("Items_Name")
            txtCost.Text = r("cost_Price")
            txtGomla.Text = r("Total_Price")
            txtCustomer.Text = r("Customer_Price")

            For Each crtl As Control In GroupBox2.Controls
                If TypeOf crtl Is TextBox AndAlso crtl.Name <> "txt_Barcode" Then
                    crtl.Text = ""
                End If
            Next
            'MsgBox(C)
        Catch ex As Exception
            Return
        End Try
    End Sub
    Private Sub Account_Cost()
        Try
            A = 0
            C = 0
            If cbo_Kind.SelectedIndex = -1 Then Return
            Myconn.ExecQuery("Select i.Parcode, iif(ISNULL(i.Customer_Price),0,i.Customer_Price) as Customer_Price,iif(ISNULL(i.cost_Price),0,i.cost_Price) as cost_Price, iif(ISNULL(i.Total_Price),0,i.Total_Price) as Total_Price, i.Supplier_ID, IIf(ISNULL(c.Pur_num - s.Sales_num), 0, (c.Pur_num - s.Sales_num)) as rest 
                                from (Items i Left join  (select iif(ISNULL(sum(Iteme_Number)),0,sum(Iteme_Number)) as Pur_num,items_cod from Purchases group by items_Cod,Status having Status = true and items_cod = " & CInt(cbo_Kind.SelectedValue) & " ) c
                                    on i.items_Cod = c.items_Cod  )
                                    left join (Select iif(ISNULL(sum(Items_num)),0,sum(Items_num)) as Sales_num ,items_Cod From Sales group by items_cod ,Status having Status = true and items_cod = " & CInt(cbo_Kind.SelectedValue) & ") S
                                    on i.items_Cod = S.items_Cod group by i.items_Cod, i.Parcode,i.Customer_Price,i.cost_Price,i.Total_Price,i.Supplier_ID,iif(ISNULL(c.Pur_num - s.Sales_num),0,(c.Pur_num - s.Sales_num))
                                    having i.items_Cod =" & CInt(cbo_Kind.SelectedValue))
            If Myconn.Recodcount = 0 Then Return
            Dim r As DataRow = Myconn.dt.Rows(0)

            A = r("rest") ' الكمية المتبقية من الصنف
            C = A * r("cost_Price") ' قيمة الكمية المتبقية من الصنف
            par = r("Parcode")
            txtAmount.Text = r("rest")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try

    End Sub

#Region "Print"
#Region "Barcode"
    Sub Print_Barcode(dgr As DataGridView)
        Try
            Dim L1 As Integer = My.Settings.B_L1
            Dim L3 As Integer = My.Settings.B_L3
            Dim L4 As Integer = My.Settings.B_L4
            Dim l5 As Integer = My.Settings.B_L5
            Dim Z As Integer = My.Settings.B_Size

            Dim table As New DataTable
            For i As Integer = 1 To 5
                Dim x As String
                x = Format(i, "00")
                table.Columns.Add(x)
            Next

            For T As Integer = 1 To My.Settings.label_number - 1
                table.Rows.Add()
                table.Rows(table.Rows.Count - 1)(0) = ""
                table.Rows(table.Rows.Count - 1)(1) = 0
                table.Rows(table.Rows.Count - 1)(2) = ""
                table.Rows(table.Rows.Count - 1)(3) = ""
            Next

            For Each dr As DataGridViewRow In drg.Rows

                For i As Integer = 0 To CInt(dr.Cells(7).Value) - 1
                    table.Rows.Add()
                    Select Case L1
                        Case 0 ' الشركة
                            table.Rows(table.Rows.Count - 1)(0) = My.Settings.Co_name
                        Case 1 '  المدير
                            table.Rows(table.Rows.Count - 1)(0) = My.Settings.Co_HR
                        Case 2 '  العنوان و التليفون
                            table.Rows(table.Rows.Count - 1)(0) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                        Case 3 'الصنف
                            table.Rows(table.Rows.Count - 1)(0) = dr.Cells(2).Value
                        Case 4 '  السعر
                            table.Rows(table.Rows.Count - 1)(0) = "السعر " & dr.Cells(10).Value & " جنيه "
                        Case 5 '    رقم الباركود
                            table.Rows(table.Rows.Count - 1)(0) = dr.Cells(3).Value
                    End Select

                    table.Rows(table.Rows.Count - 1)(1) = "*" & dr.Cells(3).Value & "*" ' الباركود
                    Select Case L3
                        Case 0 ' الشركة
                            table.Rows(table.Rows.Count - 1)(2) = My.Settings.Co_name
                        Case 1 '  المدير
                            table.Rows(table.Rows.Count - 1)(2) = My.Settings.Co_HR
                        Case 2 '  العنوان و التليفون
                            table.Rows(table.Rows.Count - 1)(2) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                        Case 3 'الصنف
                            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value
                        Case 4 '  السعر
                            table.Rows(table.Rows.Count - 1)(2) = "السعر " & dr.Cells(10).Value & " جنيه "
                        Case 5 '    رقم الباركود
                            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(3).Value
                    End Select
                    Select Case L4
                        Case 0 ' الشركة
                            table.Rows(table.Rows.Count - 1)(3) = My.Settings.Co_name
                        Case 1 '  المدير
                            table.Rows(table.Rows.Count - 1)(3) = My.Settings.Co_HR
                        Case 2 '  العنوان و التليفون
                            table.Rows(table.Rows.Count - 1)(3) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                        Case 3 'الصنف
                            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(2).Value
                        Case 4 '  السعر
                            table.Rows(table.Rows.Count - 1)(3) = "السعر " & dr.Cells(10).Value & " جنيه "
                        Case 5 '    رقم الباركود
                            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value
                    End Select
                    Select Case l5
                        Case 0 ' الشركة
                            table.Rows(table.Rows.Count - 1)(4) = My.Settings.Co_name
                        Case 1 '  المدير
                            table.Rows(table.Rows.Count - 1)(4) = My.Settings.Co_HR
                        Case 2 '  العنوان و التليفون
                            table.Rows(table.Rows.Count - 1)(4) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                        Case 3 'الصنف
                            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(2).Value
                        Case 4 '  السعر
                            table.Rows(table.Rows.Count - 1)(4) = "السعر " & dr.Cells(10).Value & " جنيه "
                        Case 5 '    رقم الباركود
                            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(3).Value
                    End Select

                    If My.Settings.B_L1_V = False Then table.Rows(table.Rows.Count - 1)(0) = ""
                    If My.Settings.B_L3_V = False Then table.Rows(table.Rows.Count - 1)(2) = ""
                    If My.Settings.B_L4_V = False Then table.Rows(table.Rows.Count - 1)(3) = ""
                    If My.Settings.B_L5_v = False Then table.Rows(table.Rows.Count - 1)(4) = ""
                Next

            Next

            Select Case Z
                Case 0 '1 × 1.8 inch - A4
                    Dim rpt As New Barcode_01
                    Dim margins As PageMargins

                    margins = rpt.PrintOptions.PageMargins
                    margins.bottomMargin = My.Settings.M_Butom
                    margins.leftMargin = My.Settings.M_Left
                    margins.rightMargin = My.Settings.M_Right
                    margins.topMargin = My.Settings.M_Top

                    rpt.PrintOptions.ApplyPageMargins(margins)

                    rpt.SetDataSource(table)
                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                        rpt.PrintToPrinter(1, False, 0, 0)
                    End If
                Case 1 '0.5 × 1.8 inch - A4
                    Dim rpt As New Barcode_02
                    Dim margins As PageMargins

                    margins = rpt.PrintOptions.PageMargins
                    margins.bottomMargin = My.Settings.M_Butom
                    margins.leftMargin = My.Settings.M_Left
                    margins.rightMargin = My.Settings.M_Right
                    margins.topMargin = My.Settings.M_Top

                    rpt.PrintOptions.ApplyPageMargins(margins)

                    rpt.SetDataSource(table)
                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                        rpt.PrintToPrinter(1, False, 0, 0)
                    End If
                Case 2 ' User(40.7 mm × 25. mm) Barcode Label
                    Dim rpt As New Barcode_03
                    rpt.SetDataSource(table)
                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                        rpt.PrintToPrinter(1, False, 0, 0)
                    End If
            'rpt.PrintToPrinter(1, False, 0, 0)
                Case 3 '2×4 ( 50.8 mm × 101.6 mm ) Barcode Label
                    Dim rpt As New Barcode_04
                    rpt.SetDataSource(table)
                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                        rpt.PrintToPrinter(1, False, 0, 0)
                    End If
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Sub Print_One_Barcode(dgr As DataGridView)
        Try
            Dim L1 As Integer = My.Settings.B_L1
            Dim L3 As Integer = My.Settings.B_L3
            Dim L4 As Integer = My.Settings.B_L4
            Dim l5 As Integer = My.Settings.B_L5
            Dim Z As Integer = My.Settings.B_Size
            Dim col As Integer = 0
            Dim table As New DataTable
            For i As Integer = 1 To 5
                Dim x As String
                x = Format(i, "00")
                table.Columns.Add(x)
            Next
            For T As Integer = 1 To My.Settings.label_number - 1
                table.Rows.Add()
                table.Rows(table.Rows.Count - 1)(0) = ""
                table.Rows(table.Rows.Count - 1)(1) = 0
                table.Rows(table.Rows.Count - 1)(2) = ""
                table.Rows(table.Rows.Count - 1)(3) = ""
            Next

            For SS As Integer = 0 To drg.SelectedRows.Count - 1
                For i As Integer = 0 To CInt(dgr(7, drg.SelectedRows(SS).Index).Value) - 1
                    table.Rows.Add()

                    Select Case L1
                        Case 0 ' الشركة
                            table.Rows(table.Rows.Count - 1)(0) = My.Settings.Co_name
                        Case 1 '  المدير
                            table.Rows(table.Rows.Count - 1)(0) = My.Settings.Co_HR
                        Case 2 '  العنوان و التليفون
                            table.Rows(table.Rows.Count - 1)(0) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                        Case 3 'الصنف
                            table.Rows(table.Rows.Count - 1)(0) = dgr(2, drg.SelectedRows(SS).Index).Value
                        Case 4 '  السعر
                            table.Rows(table.Rows.Count - 1)(0) = "السعر " & dgr(10, drg.SelectedRows(SS).Index).Value & " جنيه "
                        Case 5 '    رقم الباركود
                            table.Rows(table.Rows.Count - 1)(0) = dgr(3, drg.SelectedRows(SS).Index).Value
                    End Select

                    table.Rows(table.Rows.Count - 1)(1) = "*" & dgr(3, drg.SelectedRows(SS).Index).Value & "*" ' الباركود
                    Select Case L3
                        Case 0 ' الشركة
                            table.Rows(table.Rows.Count - 1)(2) = My.Settings.Co_name
                        Case 1 '  المدير
                            table.Rows(table.Rows.Count - 1)(2) = My.Settings.Co_HR
                        Case 2 '  العنوان و التليفون
                            table.Rows(table.Rows.Count - 1)(2) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                        Case 3 'الصنف
                            table.Rows(table.Rows.Count - 1)(2) = dgr(2, drg.SelectedRows(SS).Index).Value
                        Case 4 '  السعر
                            table.Rows(table.Rows.Count - 1)(2) = "السعر " & dgr(10, drg.SelectedRows(SS).Index).Value & " جنيه "
                        Case 5 '    رقم الباركود
                            table.Rows(table.Rows.Count - 1)(2) = dgr(3, drg.SelectedRows(SS).Index).Value
                    End Select
                    Select Case L4
                        Case 0 ' الشركة
                            table.Rows(table.Rows.Count - 1)(3) = My.Settings.Co_name
                        Case 1 '  المدير
                            table.Rows(table.Rows.Count - 1)(3) = My.Settings.Co_HR
                        Case 2 '  العنوان و التليفون
                            table.Rows(table.Rows.Count - 1)(3) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                        Case 3 'الصنف
                            table.Rows(table.Rows.Count - 1)(3) = dgr(2, drg.SelectedRows(SS).Index).Value
                        Case 4 '  السعر
                            table.Rows(table.Rows.Count - 1)(3) = "السعر " & dgr(10, drg.SelectedRows(SS).Index).Value & " جنيه "
                        Case 5 '    رقم الباركود
                            table.Rows(table.Rows.Count - 1)(3) = dgr(3, drg.SelectedRows(SS).Index).Value
                    End Select
                    Select Case l5
                        Case 0 ' الشركة
                            table.Rows(table.Rows.Count - 1)(4) = My.Settings.Co_name
                        Case 1 '  المدير
                            table.Rows(table.Rows.Count - 1)(4) = My.Settings.Co_HR
                        Case 2 '  العنوان و التليفون
                            table.Rows(table.Rows.Count - 1)(4) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                        Case 3 'الصنف
                            table.Rows(table.Rows.Count - 1)(4) = dgr(2, drg.SelectedRows(SS).Index).Value
                        Case 4 '  السعر
                            table.Rows(table.Rows.Count - 1)(4) = "السعر " & dgr(10, drg.SelectedRows(SS).Index).Value & " جنيه "
                        Case 5 '    رقم الباركود
                            table.Rows(table.Rows.Count - 1)(4) = dgr(3, drg.SelectedRows(SS).Index).Value
                    End Select

                    If My.Settings.B_L1_V = False Then table.Rows(table.Rows.Count - 1)(0) = ""
                    If My.Settings.B_L3_V = False Then table.Rows(table.Rows.Count - 1)(2) = ""
                    If My.Settings.B_L4_V = False Then table.Rows(table.Rows.Count - 1)(3) = ""
                    If My.Settings.B_L5_v = False Then table.Rows(table.Rows.Count - 1)(4) = ""
                Next

            Next
            Select Case Z
                Case 0 '1 × 1.8 inch - A4
                    Dim rpt As New Barcode_01
                    Dim margins As PageMargins

                    margins = rpt.PrintOptions.PageMargins
                    margins.bottomMargin = My.Settings.M_Butom
                    margins.leftMargin = My.Settings.M_Left
                    margins.rightMargin = My.Settings.M_Right
                    margins.topMargin = My.Settings.M_Top

                    rpt.PrintOptions.ApplyPageMargins(margins)

                    rpt.SetDataSource(table)

                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                        rpt.PrintToPrinter(1, False, 0, 0)
                    End If

                Case 1 '0.5 × 1.8 inch - A4
                    Dim rpt As New Barcode_02
                    Dim margins As PageMargins

                    margins = rpt.PrintOptions.PageMargins
                    margins.bottomMargin = My.Settings.M_Butom
                    margins.leftMargin = My.Settings.M_Left
                    margins.rightMargin = My.Settings.M_Right
                    margins.topMargin = My.Settings.M_Top

                    rpt.PrintOptions.ApplyPageMargins(margins)
                    rpt.SetDataSource(table)

                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                        rpt.PrintToPrinter(1, False, 0, 0)
                    End If

                Case 2 ' User(40.7 mm × 25. mm) Barcode Label
                    Dim rpt As New Barcode_03

                    rpt.SetDataSource(table)

                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                        rpt.PrintToPrinter(1, False, 0, 0)
                    End If
                Case 3 '2×4 ( 50.8 mm × 101.6 mm ) Barcode Label
                    Dim rpt As New Barcode_04
                    rpt.SetDataSource(table)

                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                        rpt.PrintToPrinter(1, False, 0, 0)
                    End If

            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Sub Print_Number_Barcode(dgr As DataGridView)
        Try
            Dim L1 As Integer = My.Settings.B_L1
            Dim L3 As Integer = My.Settings.B_L3
            Dim L4 As Integer = My.Settings.B_L4
            Dim l5 As Integer = My.Settings.B_L5
            Dim Z As Integer = My.Settings.B_Size
            Dim num As String = InputBox("أدخل العدد", " طباعة عدد محدد من الباركود")

            If num = Nothing Then
                Return
            End If
            Dim table As New DataTable
            For i As Integer = 1 To 5
                Dim x As String
                x = Format(i, "00")
                table.Columns.Add(x)
            Next
            For T As Integer = 1 To My.Settings.label_number - 1
                table.Rows.Add()
                table.Rows(table.Rows.Count - 1)(0) = ""
                table.Rows(table.Rows.Count - 1)(1) = 0
                table.Rows(table.Rows.Count - 1)(2) = ""
                table.Rows(table.Rows.Count - 1)(3) = ""
            Next

            For i As Integer = 1 To CInt(num)
                table.Rows.Add()
                Select Case L1
                    Case 0 ' الشركة
                        table.Rows(table.Rows.Count - 1)(0) = My.Settings.Co_name
                    Case 1 '  المدير
                        table.Rows(table.Rows.Count - 1)(0) = My.Settings.Co_HR
                    Case 2 '  العنوان و التليفون
                        table.Rows(table.Rows.Count - 1)(0) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                    Case 3 'الصنف
                        table.Rows(table.Rows.Count - 1)(0) = drg.CurrentRow.Cells(2).Value
                    Case 4 '  السعر
                        table.Rows(table.Rows.Count - 1)(0) = "السعر " & drg.CurrentRow.Cells(10).Value & " جنيه "
                    Case 5 '    رقم الباركود
                        table.Rows(table.Rows.Count - 1)(0) = drg.CurrentRow.Cells(3).Value
                End Select

                table.Rows(table.Rows.Count - 1)(1) = "*" & drg.CurrentRow.Cells(3).Value & "*" ' الباركود
                Select Case L3
                    Case 0 ' الشركة
                        table.Rows(table.Rows.Count - 1)(2) = My.Settings.Co_name
                    Case 1 '  المدير
                        table.Rows(table.Rows.Count - 1)(2) = My.Settings.Co_HR
                    Case 2 '  العنوان و التليفون
                        table.Rows(table.Rows.Count - 1)(2) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                    Case 3 'الصنف
                        table.Rows(table.Rows.Count - 1)(2) = drg.CurrentRow.Cells(2).Value
                    Case 4 '  السعر
                        table.Rows(table.Rows.Count - 1)(2) = "السعر " & drg.CurrentRow.Cells(10).Value & " جنيه "
                    Case 5 '    رقم الباركود
                        table.Rows(table.Rows.Count - 1)(2) = drg.CurrentRow.Cells(3).Value
                End Select
                Select Case L4
                    Case 0 ' الشركة
                        table.Rows(table.Rows.Count - 1)(3) = My.Settings.Co_name
                    Case 1 '  المدير
                        table.Rows(table.Rows.Count - 1)(3) = My.Settings.Co_HR
                    Case 2 '  العنوان و التليفون
                        table.Rows(table.Rows.Count - 1)(3) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                    Case 3 'الصنف
                        table.Rows(table.Rows.Count - 1)(3) = drg.CurrentRow.Cells(2).Value
                    Case 4 '  السعر
                        table.Rows(table.Rows.Count - 1)(3) = "السعر " & drg.CurrentRow.Cells(10).Value & " جنيه "
                    Case 5 '    رقم الباركود
                        table.Rows(table.Rows.Count - 1)(3) = drg.CurrentRow.Cells(3).Value
                End Select
                Select Case l5
                    Case 0 ' الشركة
                        table.Rows(table.Rows.Count - 1)(4) = My.Settings.Co_name
                    Case 1 '  المدير
                        table.Rows(table.Rows.Count - 1)(4) = My.Settings.Co_HR
                    Case 2 '  العنوان و التليفون
                        table.Rows(table.Rows.Count - 1)(4) = My.Settings.Co_address & " - " & My.Settings.Co_tel
                    Case 3 'الصنف
                        table.Rows(table.Rows.Count - 1)(4) = drg.CurrentRow.Cells(2).Value
                    Case 4 '  السعر
                        table.Rows(table.Rows.Count - 1)(4) = "السعر " & drg.CurrentRow.Cells(10).Value & " جنيه "
                    Case 5 '    رقم الباركود
                        table.Rows(table.Rows.Count - 1)(4) = drg.CurrentRow.Cells(3).Value
                End Select
                If My.Settings.B_L1_V = False Then table.Rows(table.Rows.Count - 1)(0) = ""
                If My.Settings.B_L3_V = False Then table.Rows(table.Rows.Count - 1)(2) = ""
                If My.Settings.B_L4_V = False Then table.Rows(table.Rows.Count - 1)(3) = ""
                If My.Settings.B_L5_v = False Then table.Rows(table.Rows.Count - 1)(4) = ""
            Next

            Select Case Z
                Case 0 '1 × 1.8 inch - A4
                    Dim rpt As New Barcode_01
                    Dim margins As PageMargins

                    margins = rpt.PrintOptions.PageMargins
                    margins.bottomMargin = My.Settings.M_Butom
                    margins.leftMargin = My.Settings.M_Left
                    margins.rightMargin = My.Settings.M_Right
                    margins.topMargin = My.Settings.M_Top

                    rpt.PrintOptions.ApplyPageMargins(margins)
                    rpt.SetDataSource(table)
                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                    rpt.PrintToPrinter(1, False, 0, 0)
                    End If
                Case 1 '0.5 × 1.8 inch - A4
                    Dim rpt As New Barcode_02
                    Dim margins As PageMargins

                    margins = rpt.PrintOptions.PageMargins
                    margins.bottomMargin = My.Settings.M_Butom
                    margins.leftMargin = My.Settings.M_Left
                    margins.rightMargin = My.Settings.M_Right
                    margins.topMargin = My.Settings.M_Top

                    rpt.PrintOptions.ApplyPageMargins(margins)
                    rpt.SetDataSource(table)
                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                        rpt.PrintToPrinter(1, False, 0, 0)
                    End If
                Case 2 ' User(40.0 mm × 25.0 mm) Barcode Label
                    Dim rpt As New Barcode_03

                    rpt.SetDataSource(table)
                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                        rpt.PrintToPrinter(1, False, 0, 0)
                    End If
                Case 3 '1×2 ( 40.0 mm × 12.0 mm ) Barcode Label
                    Dim rpt As New Barcode_04
                    rpt.SetDataSource(table)
                    If My.Settings.Barcode_Previo = True Then
                        frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                        frmReportViewer.Show()
                    Else
                        rpt.PrintOptions.PrinterName = My.Settings.Printer_barcode
                        rpt.PrintToPrinter(1, False, 0, 0)
                    End If
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
#End Region
    Private Sub Print_Pur_Bill()
        Dim table As New DataTable
        For i As Integer = 1 To 9
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows

            'If dr.Cells(11).Value = False Then GoTo A
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(3).Value
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(7).Value
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(8).Value
            table.Rows(table.Rows.Count - 1)(8) = dr.Cells(11).Value
            'A:
        Next
        Dim rpt As New rpt_Pur_Bill
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Bill_num", txtBill_ID.Text)
        rpt.SetParameterValue("Supplier", cbo_Supplier.Text)
        rpt.SetParameterValue("F_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("Price", Label18.Text)
        rpt.SetParameterValue("Reduce", Label18.Text - Label19.Text)
        rpt.SetParameterValue("Total", Label19.Text)
        rpt.SetParameterValue("Nisba", Label20.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub
    Private Sub Print_Pur_Bill_2()
        Dim table As New DataTable
        For i As Integer = 1 To 8
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows

            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value ' الصنف
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(3).Value ' الباركود
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(7).Value ' الكمية
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(6).Value ' التكلفة
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(4).Value ' المستهلك
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(8).Value ' إجمالي التكلفة
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(11).Value '

        Next
        Dim rpt As New rpt_Pur_Bill_2
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Bill_num", txtBill_ID.Text)
        rpt.SetParameterValue("Supplier", cbo_Supplier.Text)
        rpt.SetParameterValue("F_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("Total", Label19.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub
    Private Sub Print_bill_back()
        Dim table As New DataTable
        For i As Integer = 1 To 8
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows

            If dr.Cells(11).Value = True Then GoTo A
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(3).Value
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(7).Value
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value
            table.Rows(table.Rows.Count - 1)(5) = " % " & dr.Cells(5).Value
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(8).Value
A:
        Next
        Dim rpt As New rpt_Pur_Bill___back
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Bill_num", txtBill_ID.Text)
        rpt.SetParameterValue("Supplier", cbo_Supplier.Text)
        rpt.SetParameterValue("F_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("Price", Label18.Text)
        rpt.SetParameterValue("Reduce", Label18.Text - Label19.Text)
        rpt.SetParameterValue("Total", Label19.Text)
        rpt.SetParameterValue("Nisba", " % " & Label20.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If

    End Sub
    Private Sub Print_bill_back_2()
        Dim table As New DataTable
        For i As Integer = 1 To 8
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows

            If dr.Cells(11).Value = True Then GoTo A
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value ' الصنف
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(3).Value ' الباركود
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(7).Value ' الكمية
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(6).Value ' التكلفة
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(4).Value ' المستهلك
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(8).Value ' إجمالي التكلفة
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(11).Value '
A:
        Next
        Dim rpt As New rpt_Pur_Bill_Back_2
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Bill_num", txtBill_ID.Text)
        rpt.SetParameterValue("Supplier", cbo_Supplier.Text)
        rpt.SetParameterValue("F_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("Total", Label19.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If

    End Sub
#End Region

    Private Sub cbo_Supplier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_Supplier.SelectedIndexChanged
        ErrorProvider1.Clear()
    End Sub

    Private Sub cbo_Stock_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_Stock.SelectedIndexChanged
        ErrorProvider1.Clear()
    End Sub


End Class