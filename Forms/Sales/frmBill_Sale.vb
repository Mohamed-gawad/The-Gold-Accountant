Imports System.Globalization
Public Class frmBill_Sale
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st As String
    Dim x, y, ID, ID2 As Integer
    Dim Cost As Double
    Dim T, P, R As Double
#Region "Functions"
    Private Sub New_record()
        Try
            drg.Rows.Clear()
            'Myconn.ClearAllControls(GroupBox1, True)
            Myconn.ClearAllControls(GroupBox2, True)
            Myconn.Autonumber("Sales_Bill_ID", "Sales", txtBill_ID, Me)
            Me.Text = "فاتورة مبيعات رقم :  " & txtBill_ID.Text
            cbo_Stock.SelectedValue = My.Settings.Stock_ID
            cboPrice.SelectedIndex = My.Settings.Sales_Kind
            cbo_Customer.SelectedValue = 1
            If My.Settings.Reduce = True Then
                txt_Reduce.Text = My.Settings.Reduce_amount
            Else
                txt_Reduce.Text = ""
            End If
            Label18.Text = 0
            Label19.Text = 0
            Label20.Text = 0
            Label22.Text = 0
            Label23.Text = 0
            txtReduce.Text = 0

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")

        End Try
    End Sub
    Private Sub Filldrg()
        Try
            drg.Rows.Clear()
            Select Case y
                Case 0
                    st = " where Sales.Sales_Bill_ID =" & CInt(txtBill_ID.Text)
                Case 1
                    st = " where Sales.Sales_Bill_ID =" & CInt(txtSearch.Text)
            End Select
            Myconn.ExecQuery("SELECT Sales.Sales_Date, Sales.Sales_time, Sales.Sales_Bill_ID, Sales.items_Cod, Sales.Items_Price,Stock_Name, Sales.Items_num, Sales.Total_Price, Sales.Reduce, Sales.Final_Total_Price, Sales.Status, Sales.ID, Sales.Bill_Kind, Sales.Sales_Kind_ID, Items.Items_Name, Employees.Employee_Name, Items.Parcode
                                FROM (Items RIGHT JOIN (Sales LEFT JOIN (Users_ID LEFT JOIN Employees ON Users_ID.Employee_ID = Employees.Employee_ID)
                                ON Sales.Users_ID = Users_ID.Employee_ID) ON Items.items_Cod = Sales.items_Cod) LEFT JOIN Stocks ON Sales.Stock_ID = Stocks.Stock_ID
                                " & st & " order by Sales.ID ")

            If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception, MsgBoxStyle.Critical, "رسالة") : Exit Sub
            Dim V1 As Double = 0
            Dim V2 As Double = 0
            Dim B As Double = 0
            For i As Integer = 0 To Myconn.dt.Rows.Count - 1
                Dim r As DataRow = Myconn.dt.Rows(i)
                drg.Rows.Add()
                drg.Rows(i).Cells(0).Value = i + 1
                drg.Rows(i).Cells(1).Value = r("Sales_time")
                drg.Rows(i).Cells(2).Value = r("Items_Name")
                drg.Rows(i).Cells(3).Value = r("Parcode")
                drg.Rows(i).Cells(4).Value = r("Items_num")
                drg.Rows(i).Cells(5).Value = r("Items_Price")
                drg.Rows(i).Cells(6).Value = r("Reduce")
                drg.Rows(i).Cells(7).Value = r("Final_Total_Price")
                drg.Rows(i).Cells(8).Value = r("Employee_Name")
                drg.Rows(i).Cells(9).Value = r("Status")
                drg.Rows(i).Cells(10).Value = r("ID")
                drg.Rows(i).Cells(11).Value = r("Stock_Name")

                If r("Sales_Kind_ID").ToString = False Then
                    drg.Rows(i).Cells(2).Style.BackColor = Color.Yellow
                Else
                    drg.Rows(i).Cells(2).Style.BackColor = Color.LemonChiffon
                End If

                If r("Bill_Kind").ToString = False Then
                    drg.Rows(i).Cells(5).Style.BackColor = Color.Pink
                Else
                    drg.Rows(i).Cells(5).Style.BackColor = Color.LemonChiffon
                End If

                If drg.Rows(i).Cells(9).Value = True Then
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.LemonChiffon
                    V1 += CDec(drg.Rows(i).Cells(4).Value) * drg.Rows(i).Cells(5).Value
                    V2 += CDec(drg.Rows(i).Cells(7).Value)
                Else
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    drg.Rows(i).Cells(2).Style.BackColor = Color.Red
                    drg.Rows(i).Cells(5).Style.BackColor = Color.Red
                    B += CDec(drg.Rows(i).Cells(7).Value)
                End If
            Next
            Myconn.DataGridview_MoveLast(drg, 0)
            Label19.Text = Math.Round((V2 - Val(txtReduce.Text)), 2)
            Label18.Text = Math.Round(V1, 2)
            Label22.Text = Math.Round(B, 2)
            Label23.Text = Math.Round((V1 - V2), 2)
            Label20.Text = "%" & Math.Round(((Val(Val(Label18.Text) - Val(Label19.Text)) / Val(Label18.Text)) * 100), 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Binding()
        Try

            Select Case x
                Case 0
                    Myconn.ExecQuery("Select i.Parcode, i.Customer_Price, i.cost_Price, i.Total_Price, i.Supplier_ID, IIf(ISNULL((c.Pur_num + T.amount_t) - (s.Sales_num + F.amount_F)), 0, ((c.Pur_num + t.amount_t) - (s.Sales_num + amount_F))) as rest  
                                   from (((Items i Left join  (select iif(ISNULL(sum(Iteme_Number)),0,sum(Iteme_Number)) as Pur_num,items_cod from Purchases group by items_Cod,Status,Stock_ID having Status = True and Stock_ID = " & CInt(cbo_Stock.SelectedValue) & "  ) c
                                    on i.items_Cod = c.items_Cod  )
                                    left join (Select iif(ISNULL(sum(Items_num)),0,sum(Items_num)) as Sales_num ,items_Cod From Sales group by items_cod,Status,Stock_ID having  Status = True and Stock_ID = " & CInt(cbo_Stock.SelectedValue) & "  ) S
                                    on i.items_Cod = S.items_Cod )
                                    left join  (Select iif(Isnull(Sum(items_amount)),0,Sum(items_amount)) as amount_F, items_cod From Items_move  group by items_cod,Stock_From having Stock_From = " & CInt(cbo_Stock.SelectedValue) & " ) F
                                    on  i.items_Cod = F.items_cod)
                                    left join  (Select iif(Isnull(Sum(items_amount)),0,Sum(items_amount)) as amount_t, items_cod From Items_move  group by items_cod,Stock_to having Stock_to = " & CInt(cbo_Stock.SelectedValue) & " ) T
                                    on  i.items_Cod = T.items_cod
                                    group by i.items_Cod, i.Parcode,i.Customer_Price,i.cost_Price,i.Total_Price,i.Supplier_ID,IIf(ISNULL((c.Pur_num + T.amount_t) - (s.Sales_num + F.amount_F)), 0, ((c.Pur_num + t.amount_t) - (s.Sales_num + amount_F)))
                                    having i.items_Cod =" & CInt(cbo_Kind.SelectedValue))

                    Dim r As DataRow = Myconn.dt.Rows(0)
                    txtAmount_Kind.Text = If(IsDBNull(r("rest")), 0, r("rest"))
                    txtPrice_Customer.Text = If(cboPrice.SelectedIndex = 0, If(IsDBNull(r("Customer_Price")), 0, r("Customer_Price")), If(IsDBNull(r("Total_Price")), 0, r("Total_Price")))
                    txt_Barcode.Text = r("Parcode")
                    Cost = If(IsDBNull(r("cost_Price")), 0, r("cost_Price"))

                Case 1
                    Myconn.ExecQuery("Select S.Sales_Date,S.Sales_Bill_ID,S.Customer_ID,S.Stock_ID,S.Bill_Kind,S.Sales_Kind_ID,P.Amount from Sales S 
                                        Left join Safe_payment_per P on S.Sales_Bill_ID = P.Sales_ID
                                        where S.Sales_Bill_ID =" & CInt(txtSearch.Text))

                    If Myconn.dt.Rows.Count = 0 Then MsgBox(" هذه الفاتورة غير موجودة", MsgBoxStyle.Critical & MsgBoxStyle.MsgBoxRtlReading, "رسالة") : Return
                    Dim r As DataRow = Myconn.dt.Rows(0)
                    D_date.Text = r("Sales_Date")
                    txtBill_ID.Text = r("Sales_Bill_ID")
                    cbo_Stock.SelectedValue = r("Stock_ID")
                    cbo_Customer.SelectedValue = r("Customer_ID")
                    R1.Checked = If(r("Sales_Kind_ID").ToString = True, True, False)
                    R2.Checked = If(r("Sales_Kind_ID").ToString = True, False, True)
                    cboPrice.SelectedIndex = If(r("Bill_Kind").ToString = True, 0, 1)
                    txtReduce.Text = If(IsDBNull(r("Amount")), 0, r("Amount"))
                    Me.Text = "فاتورة مبيعات رقم :  " & r("Sales_Bill_ID")
                Case 2
                    Myconn.ExecQuery("Select * from Sales where ID =" & CInt(drg.CurrentRow.Cells(10).Value))
                    If Myconn.Recodcount = 0 Then Return
                    Dim r As DataRow = Myconn.dt.Rows(0)
                    cbo_Kind.SelectedValue = r("items_Cod")
                    cboPrice.SelectedIndex = If(r("Bill_Kind").ToString = True, 0, 1)
                    txt_Amount.Text = r("Items_num")
                    txt_Reduce.Text = r("Reduce")
                    cbo_Stock.SelectedValue = r("Stock_ID")
                    R1.Checked = If(r("Sales_Kind_ID").ToString = True, True, False)
                    R2.Checked = If(r("Sales_Kind_ID").ToString = True, False, True)
            End Select

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Save_recod()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@Sales_Date", Format(CDate(D_date.Text), "yyyy/MM/dd")) ' التاريخ
                .Addparam("@Sales_time", Label12.Text) ' الوقت
                .Addparam("@Sales_Bill_ID", txtBill_ID.Text) ' رقم الفاتورة
                .Addparam("@items_Cod", cbo_Kind.SelectedValue) ' الصنف
                .Addparam("@Items_Price", txtPrice_Customer.Text) ' سعر المستهلك
                .Addparam("@Items_num", txt_Amount.Text) ' الكمية
                .Addparam("@Total_Price", Val(txtPrice_Customer.Text) * Val(txt_Amount.Text)) ' الاجمالي قبل الخصم
                .Addparam("@Reduce", txt_Reduce.Text) ' الخصم
                .Addparam("@Final_Total_Price", txt_Total.Text) ' الاجمالي بعد الخصم
                .Addparam("@Customer_ID", cbo_Customer.SelectedValue) ' العميل
                .Addparam("@Users_ID", My.Settings.user_ID) ' المستخدم
                .Addparam("@Sales_Kind_ID", If(R1.Checked = True, 1, 0)) ' نقدي و آجل
                .Addparam("@Stock_ID", cbo_Stock.SelectedValue) ' المخزن
                .Addparam("@Status", 1) ' الحالة
                .Addparam("@Earning", Math.Round((Val(txt_Total.Text) - Val(Val(txt_Amount.Text) * Cost)), 2)) ' الربح
                .Addparam("@Bill_Kind", If(cboPrice.SelectedIndex = 0, 1, 0)) ' جملة و قطاعي

                .ExecQuery("insert into [Sales] (Sales_Date,Sales_time,Sales_Bill_ID,items_Cod,Items_Price,Items_num,Total_Price,Reduce,Final_Total_Price,Customer_ID,Users_ID,Sales_Kind_ID,Stock_ID,Status,Earning,Bill_Kind)
                                      values(@Sales_Date,@Sales_time,@Sales_Bill_ID,@items_Cod,@Items_Price,@Items_num,@Total_Price,@Reduce,@Final_Total_Price,@Customer_ID,@Users_ID,@Sales_Kind_ID,@Stock_ID,@Status,@Earning,@Bill_Kind)")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Save_record_recive()
        Try
            ID = 0
            Myconn.Autonumber("per_ID", "Safe_Recive_per", txt1, Me)
            Myconn.ExecQuery("Select max(ID) as ID from Sales")
            Dim r As DataRow = Myconn.dt.Rows(0)
            ID = r("ID")
            With Myconn
                .Parames.Clear()
                .Addparam("@per_ID", txt1.Text)
                .Addparam("@per_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
                .Addparam("@per_time", Label12.Text)
                .Addparam("@users_ID", My.Settings.user_ID)
                .Addparam("@Recive_Item_ID", 1)
                .Addparam("@perm_ID", 1)
                .Addparam("@Amount", txt_Total.Text)
                .Addparam("@Customer_ID", cbo_Customer.SelectedValue)
                .Addparam("@Note_per", DBNull.Value)
                .Addparam("@Status", 1)
                .Addparam("@Sales_ID", ID)
                .ExecQuery("insert into [Safe_Recive_per] (per_ID,per_date,per_time,users_ID,Recive_Item_ID,perm_ID,Amount,Customer_ID,Note_per,Status,Sales_ID) Values(@per_ID,@per_date,@per_time,@users_ID,@Recive_Item_ID,@perm_ID,@Amount,@Customer_ID,@Note_per,@Status,@Sales_ID)")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Save_recod_Payment()
        Try
            ID2 = 0
            Myconn.Autonumber("per_ID", "Safe_payment_per", txt2, Me)

            With Myconn
                .Parames.Clear()
                .Addparam("@per_ID", txt2.Text)
                .Addparam("@per_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
                .Addparam("@per_time", Label12.Text)
                .Addparam("@users_ID", My.Settings.user_ID)
                .Addparam("@pay_Item_ID", 5)
                .Addparam("@perm_ID", 2)
                .Addparam("@Amount", txtReduce.Text)
                .Addparam("@Supplier_ID", DBNull.Value)
                .Addparam("@Note_per", " خصم إضافي على الفاتورة رقم " & txtBill_ID.Text)
                .Addparam("@Status", 1)
                .Addparam("@Sales_ID", txtBill_ID.Text)
                .ExecQuery("insert into [Safe_payment_per] (per_ID,per_date,per_time,users_ID,pay_Item_ID,perm_ID,Amount,Supplier_ID,Note_per,Status,Sales_ID) Values(@per_ID,@per_date,@per_time,@users_ID,@pay_Item_ID,@perm_ID,@Amount,@Supplier_ID,@Note_per,@Status,@Sales_ID)")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
        MessageBox.Show("تمت تسجيل الخصم بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub
    Private Sub Update_record()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@Sales_Date", Format(CDate(D_date.Text), "yyyy/MM/dd")) ' التاريخ
                .Addparam("@Sales_time", Label12.Text) ' الوقت
                .Addparam("@Sales_Bill_ID", txtBill_ID.Text) ' رقم الفاتورة
                .Addparam("@items_Cod", cbo_Kind.SelectedValue) ' الصنف
                .Addparam("@Items_Price", txtPrice_Customer.Text) ' سعر المستهلك
                .Addparam("@Items_num", txt_Amount.Text) ' الكمية
                .Addparam("@Total_Price", Val(txtPrice_Customer.Text) * Val(txt_Amount.Text)) ' الاجمالي قبل الخصم
                .Addparam("@Reduce", txt_Reduce.Text) ' الخصم
                .Addparam("@Final_Total_Price", txt_Total.Text) ' الاجمالي بعد الخصم
                .Addparam("@Customer_ID", cbo_Customer.SelectedValue) ' العميل
                .Addparam("@Users_ID", My.Settings.user_ID) ' المستخدم
                .Addparam("@Sales_Kind_ID", If(R1.Checked = True, 1, 0)) ' نقدي و آجل
                .Addparam("@Stock_ID", cbo_Stock.SelectedValue) ' المخزن
                '.Addparam("@Status", 1) ' الحالة
                .Addparam("@Earning", Math.Round((Val(txt_Total.Text) - Val(Val(txt_Amount.Text) * Cost)), 2)) ' الربح
                .Addparam("@Bill_Kind", If(cboPrice.SelectedIndex = 0, 1, 0)) ' جملة و قطاعي
                .Addparam("@ID", drg.CurrentRow.Cells(10).Value)
                .ExecQuery("Update [Sales] set Sales_Date=@Sales_Date,Sales_time=@Sales_time,Sales_Bill_ID=@Sales_Bill_ID,items_Cod=@items_Cod,Items_Price=@Items_Price,Items_num=@Items_num,Total_Price=@Total_Price,
                                               Reduce=@Reduce,Final_Total_Price=@Final_Total_Price,Customer_ID=@Customer_ID,Users_ID=@Users_ID,Sales_Kind_ID=@Sales_Kind_ID,Stock_ID=@Stock_ID,Earning=@Earning,Bill_Kind=@Bill_Kind where ID = @ID ")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub update_record_recive()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@per_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
                .Addparam("@per_time", Label12.Text)
                .Addparam("@users_ID", My.Settings.user_ID)
                .Addparam("@Recive_Item_ID", 1)
                .Addparam("@perm_ID", 1)
                .Addparam("@Amount", txt_Total.Text)
                .Addparam("@Customer_ID", cbo_Customer.SelectedValue)
                .Addparam("@Note_per", DBNull.Value)
                .Addparam("@Sales_ID", drg.CurrentRow.Cells(10).Value)

                .ExecQuery("Update [Safe_Recive_per]  Set per_date=@per_date,per_time=@per_time,users_ID=@users_ID,Recive_Item_ID=@Recive_Item_ID,perm_ID=@perm_ID,Amount=@Amount,Customer_ID=@Customer_ID,Note_per=@Note_per where Sales_ID =@Sales_ID")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Update_record_Payment()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@per_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
                .Addparam("@per_time", Label12.Text)
                .Addparam("@users_ID", My.Settings.user_ID)
                .Addparam("@pay_Item_ID", 5)
                .Addparam("@Amount", txtReduce.Text)
                .Addparam("@Supplier_ID", If(cbo_Customer.SelectedIndex = -1, DBNull.Value, cbo_Customer.SelectedValue))
                .Addparam("@Note_per", " خصم إضافي على الفاتورة رقم " & txtBill_ID.Text)
                .Addparam("@Sales_ID", txtBill_ID.Text)

                .ExecQuery("Update [Safe_payment_per]  Set per_date=@per_date,per_time=@per_time,users_ID=@users_ID,pay_Item_ID=@pay_Item_ID,Amount=@Amount,Supplier_ID=@Supplier_ID,Note_per=@Note_per where Sales_ID =@Sales_ID")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
        MessageBox.Show("تمت عملية تعديل الخصم بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub
    Private Sub Later_acc()
        Myconn.ExecQuery("SELECT iif(isnull(Sum(Final_Total_Price)),0,Sum(Final_Total_Price)) as Total, Customer_ID,Sales_Date From Sales
                            group by Customer_ID,Status,Sales_Date having  Status = True and Customer_ID =" & CInt(cbo_Customer.SelectedValue) & " and  Sales_Date <= #" & Format(CDate(D_date.Text), "yyyy/MM/dd") & "#")
        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        T = 0
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            T += r("Total")
        Next

        Myconn.ExecQuery("Select iif(isnull(Sum(Amount)),0,Sum(Amount)) as Amount2,Customer_ID,per_date From Safe_Recive_per 
                            group by Customer_ID,Status,per_date having  Status = True and Customer_ID =" & CInt(cbo_Customer.SelectedValue) & " and  per_date <= #" & Format(CDate(D_date.Text), "yyyy/MM/dd") & "#")
        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        P = 0
        For S As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(S)
            P += r("Amount2")
        Next

        R = Math.Round((T - P) - Val(Label19.Text), 2)
    End Sub
#End Region
    Private Sub frmBill_Sale_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        txtPrice_Customer.ReadOnly = My.Settings.Price_sales
        txt_Reduce.ReadOnly = My.Settings.Reduce
        txt_Reduce.Text = My.Settings.Reduce_amount
        cboPrice.SelectedIndex = My.Settings.Sales_Kind
        If My.Settings.Sales_Case = False Then
            cboPrice.Enabled = False
        Else
            cboPrice.Enabled = True
        End If
        cbo_Stock.SelectedValue = My.Settings.Stock_ID
        If My.Settings.S_Stock = False Then
            cbo_Stock.Enabled = False
        Else
            cbo_Stock.Enabled = True
        End If
        txt_Barcode.Focus()
    End Sub
    Private Sub frmBill_Sale_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width

            If F <> 1 Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
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
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
        fin = False
        Myconn.Fillcombo("Select Items_Name, items_Cod from [items] order by [Items_Name]", "[items]", "items_Cod", "Items_Name", Me, cbo_Kind)
        Myconn.Fillcombo("select * from Customers order by Customer_Name", "Customers", "Customer_ID", "Customer_Name", Me, cbo_Customer)
        Myconn.Fillcombo("select * from [Stocks] order by Stock_ID", "[Stocks]", "Stock_ID", "Stock_Name", Me, cbo_Stock)
        fin = True
        Timer1.Start()
        R1.Checked = True

        New_record()
        txtPrice_Customer.ReadOnly = My.Settings.Price_sales
        txt_Reduce.ReadOnly = My.Settings.Reduce
        txt_Reduce.Text = My.Settings.Reduce_amount
        cboPrice.SelectedIndex = My.Settings.Sales_Kind
        If My.Settings.Sales_Case = False Then
            cboPrice.Enabled = False
        End If
        cbo_Stock.SelectedValue = My.Settings.Stock_ID
        If My.Settings.S_Stock = False Then
            cbo_Stock.Enabled = False
        End If
        txt_Barcode.Focus()

        '  -------------------------------------------------------------------------------------------------- النسخة التجريبية
        'Myconn.ExecQuery("select * from Sales")
        'If Myconn.Recodcount > 200 Then
        '    MsgBox("هذه النسخة تجريبية", MsgBoxStyle.Critical, "رسالة")
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

            If Val(txt_Amount.Text) > Val(txtAmount_Kind.Text) Then
                MsgBox("الكمية لا تسمح", MsgBoxStyle.Critical, "رسالة")
                txt_Amount.Focus()

                Return
            End If
            If Val(txtPrice_Customer.Text) < Cost Then
                MsgBox("سعر البيع غير صحيح", MsgBoxStyle.Critical, "رسالة")
                txtPrice_Customer.Focus()

                Return
            End If
            Save_recod()
            If R1.Checked = True Then ' إذا كانت عملية البيع نقدية فان الاذن يذهب الى الخزنة مباشرة
                Save_record_recive()
            End If

            y = 0
            Filldrg()

            For Each txt As Control In GroupBox2.Controls
                If TypeOf txt Is TextBox Then
                    If My.Settings.Reduce = True AndAlso txt.Name = "txt_Reduce" Then
                        txt.Text = My.Settings.Reduce_amount
                    Else
                        txt.Text = ""
                    End If

                End If
            Next
            x = 0
            Binding()
            'txt_Barcode.Focus()
            MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

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

            If Val(txt_Amount.Text) > Val(txtAmount_Kind.Text) Then
                MsgBox("الكمية لا تسمح", MsgBoxStyle.Critical, "رسالة")
                txt_Amount.Focus()

                Return
            End If
            If Val(txtPrice_Customer.Text) < Cost Then
                MsgBox("سعر البيع غير صحيح", MsgBoxStyle.Critical, "رسالة")
                txtPrice_Customer.Focus()

                Return
            End If
            update_record_recive() ' تعديل إذن استلام النقدية
            Update_record() ' تعديل المبيعات
            MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

            y = 0
            Filldrg()
            x = 0
            Binding()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        Try
            If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub

            With Myconn
                .Addparam("@ID", drg.CurrentRow.Cells(10).Value)
                .ExecQuery("delete from [Safe_Recive_per] where Sales_ID = @ID") ' الحذف من جدول اذونات استلام النقدية
            End With
            If Myconn.NoErrors(True) = False Then Exit Sub
            With Myconn
                .Addparam("@ID", drg.CurrentRow.Cells(10).Value)
                .ExecQuery("delete from [Sales] where ID = @ID ") ' الحذف من المبيعات
            End With
            If drg.Rows.Count = 1 Then
                With Myconn
                    .Addparam("@ID", txtBill_ID.Text)
                    .ExecQuery("delete from [Safe_payment_per] where Sales_ID = @ID") ' حذف الخصومات الاضافية  من جدول المدفوعات
                End With
                If Myconn.NoErrors(True) = False Then Exit Sub
            End If
            If Myconn.NoErrors(True) = False Then Exit Sub
            drg.Rows.Remove(drg.CurrentRow)
            'Myconn.ClearAllControls(GroupBox2, True)

            y = 0
            Filldrg()
            For Each txt As Control In GroupBox2.Controls
                If TypeOf txt Is TextBox Then
                    If My.Settings.Reduce = True AndAlso txt.Name = "txt_Reduce" Then
                        txt.Text = My.Settings.Reduce_amount
                    Else
                        txt.Text = ""
                    End If

                End If
            Next
            If My.Settings.Reduce = True Then
                txt_Reduce.Text = My.Settings.Reduce_amount
            Else
                txt_Reduce.Text = ""
            End If
            x = 0
            Binding()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try

    End Sub
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Try
            If drg.CurrentRow.Cells(9).Value = True Then
                With Myconn
                    .Parames.Clear()
                    .Addparam("@Status", False)
                    .Addparam("@ID", drg.CurrentRow.Cells(10).Value)
                End With

            Else
                With Myconn
                    .Parames.Clear()
                    .Addparam("@Status", True)
                    .Addparam("@ID", drg.CurrentRow.Cells(10).Value)
                End With

            End If
            Myconn.ExecQuery(" Update  Sales set Status = @Status where ID = @ID") ' المبيعات
            '--------------------------------------------------------------------------------------------------------------
            If drg.CurrentRow.Cells(9).Value = True Then
                With Myconn
                    .Parames.Clear()
                    .Addparam("@Status", False)
                    .Addparam("@ID", drg.CurrentRow.Cells(10).Value)
                End With

            Else
                With Myconn
                    .Parames.Clear()
                    .Addparam("@Status", True)
                    .Addparam("@ID", drg.CurrentRow.Cells(10).Value)
                End With

            End If
            Myconn.ExecQuery(" Update  Safe_Recive_per set Status = @Status where Sales_ID = @ID") ' أذونات الاستلام من الخزنة
            y = 0
            Filldrg()
            x = 0
            Binding()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try

    End Sub
    Public Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try
            If txtSearch.Text = "" Then Return
            x = 1
            Binding()
            y = 1
            Filldrg()
            Later_acc()
            Label25.Text = R
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label19.Text = Math.Round((Val(Label19.Text) - Val(txtReduce.Text)), 2)
        Myconn.ExecQuery("Select Sales_ID from Safe_payment_per where Sales_ID = " & CInt(txtBill_ID.Text))
        If Myconn.Recodcount = 0 Then
            If txtReduce.Text = 0 Or txtReduce.Text = "" Then
                MsgBox("أدخل الخصم الإضافي", MsgBoxStyle.Critical, "رسالة")
                Return
            End If
            Save_recod_Payment()
        Else
            Update_record_Payment()
        End If
    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        If My.Settings.Sales_Bill_Kind = 0 Then
            Print_drg_Kasheer(drg)
        Else
            Print_drg_A4(drg)
        End If
    End Sub
#End Region
#Region "ComboBoxs"
    Private Sub cbo_Kind_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_Kind.SelectedIndexChanged, cboPrice.SelectedIndexChanged
        Try
            If Not fin Then Return
            If cbo_Kind.SelectedIndex = -1 Then
                txt_Barcode.Text = Nothing
                txtAmount_Kind.Text = Nothing
                txtPrice_Customer.Text = Nothing
                Return
            End If
            If cbo_Stock.SelectedIndex = -1 Then MsgBox("قم باختيار المخزن الذي سيتم البيع منه", MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة") : Return
            txt_Amount.Text = ""
            If My.Settings.Reduce = False Then txt_Reduce.Text = ""
            txt_Total.Text = ""
            x = 0
            Binding()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub cbo_Stock_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_Stock.SelectedIndexChanged
        Try
            If Not fin Then Return
            If cbo_Kind.SelectedIndex = -1 Then Return
            x = 0
            Binding()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub cbo_Customer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_Customer.SelectedIndexChanged
        If Not fin Then Return
        Later_acc()
        Label25.Text = R
    End Sub
    Private Sub cboPrice_Enter(sender As Object, e As EventArgs) Handles cboPrice.Enter
        Myconn.langAR()
    End Sub
    Private Sub cbo_Customer_Enter(sender As Object, e As EventArgs) Handles cbo_Customer.Enter
        Myconn.langAR()
    End Sub
#End Region
#Region "TextBoxs"
    Private Sub txtPrice_Customer_TextChanged(sender As Object, e As EventArgs) Handles txtPrice_Customer.TextChanged
        txt_Amount.Text = Nothing
        txt_Total.Text = Nothing
    End Sub
    Private Sub txt_Amount_TextChanged(sender As Object, e As EventArgs) Handles txt_Amount.TextChanged
        Try
            ErrorProvider1.Clear()
            If txt_Reduce.Text = Nothing OrElse txt_Amount.Text = Nothing Then Return
            Dim T As Double = Val(txt_Amount.Text) * Val(txtPrice_Customer.Text)
            Dim N As Double = T * Val(txt_Reduce.Text) / 100
            txt_Total.Text = Math.Round((T - N), 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub txt_Reduce_TextChanged(sender As Object, e As EventArgs) Handles txt_Reduce.TextChanged
        Try
            ErrorProvider1.Clear()
            If txt_Reduce.Text = Nothing OrElse txt_Amount.Text = Nothing Then Return
            Dim T As Double = Val(txt_Amount.Text) * Val(txtPrice_Customer.Text)
            Dim N As Double = T * Val(txt_Reduce.Text) / 100
            txt_Total.Text = Math.Round((T - N), 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
#Region "Moving" ' حركة المفاتيح
    Private Sub txt_Barcode_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_Barcode.KeyUp
        Try
            If txt_Barcode.Text = "" Then Return
            If cbo_Stock.SelectedIndex = -1 Then MsgBox("قم باختيار المخزن الذي سيتم البيع منه", MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة") : Return
            If e.KeyCode = Keys.Enter = True Then
                Myconn.ExecQuery("Select items_Cod from items where Parcode Like '" & txt_Barcode.Text & "'")
        If Myconn.dt.Rows.Count = 0 Then
                    cbo_Kind.SelectedIndex = -1
                    MsgBox("الصنف غير موجود أو الباركود غير صحيح", MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة")

                    Return
                End If
                Dim r As DataRow = Myconn.dt.Rows(0)
                cbo_Kind.SelectedValue = If(IsDBNull(r("items_Cod")), -1, r("items_Cod"))
                If txt_Reduce.ReadOnly = True Then
                    txt_Amount.Focus()
                    txt_Amount.Text = 1
                    txt_Amount.SelectAll()

                Else
                    txt_Reduce.Focus()
                    txt_Reduce.SelectAll()

                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
            Return
        End Try
    End Sub
    Private Sub txt_Reduce_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_Reduce.KeyUp
        If My.Settings.Reduce = True Then Return
        If e.KeyCode = Keys.Enter Then
            txt_Amount.Focus()
            txt_Amount.Text = 1
            txt_Amount.SelectAll()
        End If
    End Sub
    Private Sub txt_Amount_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_Amount.KeyUp
        If e.KeyCode = Keys.Enter Then
            If btnSave.Enabled = False Then MsgBox("هذه النسخة تجريبية", MsgBoxStyle.Critical, "رسالة") : Return
            If txt_Amount.Text = Nothing Then
                MsgBox("أدخل الكيمة", MsgBoxStyle.Critical, "رسالة")
                txt_Amount.Focus()
            Else
                btnSave_Click(Nothing, Nothing)
                If Val(txt_Amount.Text) > Val(txtAmount_Kind.Text) Then Return
                txt_Barcode.Focus()
            End If
        End If
    End Sub
    Private Sub cbo_Kind_KeyUp(sender As Object, e As KeyEventArgs) Handles cbo_Kind.KeyUp
        If e.KeyCode = Keys.Enter = True Then
            If txt_Reduce.ReadOnly = True Then
                txt_Amount.Focus()
                txt_Amount.Text = 1
                txt_Amount.SelectAll()
            Else
                txt_Reduce.Focus()
            End If
        End If
    End Sub
#End Region

    Private Sub txt_Amount_Enter(sender As Object, e As EventArgs) Handles txt_Amount.Enter
        txt_Amount.Text = Nothing
    End Sub
    Private Sub txt_Reduce_Enter(sender As Object, e As EventArgs) Handles txt_Reduce.Enter
        If My.Settings.Reduce = True Then Return
        txt_Reduce.Text = Nothing
    End Sub
    Private Sub txt_Barcode_Enter(sender As Object, e As EventArgs) Handles txt_Barcode.Enter
        txt_Barcode.Text = Nothing
        cbo_Kind.SelectedIndex = -1
        txtAmount_Kind.Text = Nothing
        txtPrice_Customer.Text = Nothing
    End Sub
    Private Sub txt_Amount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_Amount.KeyPress
        Myconn.NumberOnly(txt_Amount, e)
    End Sub
    Private Sub txtPrice_Customer_Enter(sender As Object, e As EventArgs) Handles txtPrice_Customer.Enter
        If My.Settings.Price_sales = True Then Return
        txtPrice_Customer.Text = Nothing
    End Sub
    Private Sub txtPrice_Customer_KeyUp(sender As Object, e As KeyEventArgs) Handles txtPrice_Customer.KeyUp
        If e.KeyCode = Keys.Enter Then
            If My.Settings.Reduce = True Then
                txt_Amount.Focus()
            Else
                txtReduce.Focus()
            End If
        End If
    End Sub
    Private Sub txtPrice_Customer_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPrice_Customer.KeyPress
        Myconn.NumberOnly(txtPrice_Customer, e)
    End Sub
    Private Sub txt_Reduce_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_Reduce.KeyPress
        Myconn.NumberOnly(txt_Reduce, e)
    End Sub
    Private Sub txtSearch_KeyUp(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyUp
        If e.KeyCode = Keys.Enter Then
            btnSearch_Click(Nothing, Nothing)
        End If
    End Sub
#End Region
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label12.Text = TimeOfDay.ToString("hh:mm:ss tt", CultureInfo.CreateSpecificCulture("ar-eg"))
    End Sub
    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Try
            x = 2
            Binding()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Sub Print_drg_Kasheer(dgr As DataGridView) ' طباعة الفاتورة على ورق كاشيير
        Try
            Dim rpt As New rpt_bill
            Dim table As New DataTable
            For i As Integer = 1 To 4
                Dim x As String
                x = Format(i, "00")
                table.Columns.Add(x)
            Next

            For Each dr As DataGridViewRow In drg.Rows
                If dr.Cells(9).Value = False Then GoTo a
                table.Rows.Add()
                table.Rows(table.Rows.Count - 1)(0) = dr.Cells(2).Value
                table.Rows(table.Rows.Count - 1)(1) = dr.Cells(4).Value
                table.Rows(table.Rows.Count - 1)(2) = dr.Cells(5).Value
                table.Rows(table.Rows.Count - 1)(3) = Math.Round((Val(dr.Cells(4).Value) * Val(dr.Cells(5).Value)), 2)
a:
            Next
            rpt.SetDataSource(table)
            rpt.SetParameterValue("Co_name", My.Settings.Co_name)
            rpt.SetParameterValue("Address", My.Settings.Co_address & " ت : " & My.Settings.Co_tel)
            rpt.SetParameterValue("Bill", txtBill_ID.Text)
            rpt.SetParameterValue("B_Date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
            rpt.SetParameterValue("Customer", cbo_Customer.Text)
            rpt.SetParameterValue("Total", Label18.Text)
            rpt.SetParameterValue("Reduce", Math.Round((Val(Label23.Text) + Val(txtReduce.Text)), 2))
            rpt.SetParameterValue("Total_pur", Label19.Text)
            rpt.SetParameterValue("HR", My.Settings.Co_HR)
            rpt.SetParameterValue("Later_acc", Label25.Text)

            If My.Settings.Print = True Then ' في حالة المعاينة
                rpt.PrintOptions.PrinterName = My.Settings.Printer_Sales
                frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                frmReportViewer.Show()
            Else ' في حالة عدم المعاينة والطباعة مباشرة
                rpt.PrintOptions.PrinterName = My.Settings.Printer_Sales ' طابعة فواتير المبيعات
                rpt.PrintToPrinter(1, False, 0, 0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Sub Print_drg_A4(dgr As DataGridView) ' طباعة الفاتورة على ورق A4
        Try
            Dim rpt As New rpt_bill_A4
            Dim table As New DataTable
            For i As Integer = 1 To 8
                Dim x As String
                x = Format(i, "00")
                table.Columns.Add(x)
            Next

            For Each dr As DataGridViewRow In drg.Rows
                If dr.Cells(9).Value = False Then GoTo a
                table.Rows.Add()
                table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count ' المسلسل
                table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value ' الصنف
                table.Rows(table.Rows.Count - 1)(2) = dr.Cells(3).Value ' الباركود
                table.Rows(table.Rows.Count - 1)(3) = dr.Cells(4).Value 'العدد
                table.Rows(table.Rows.Count - 1)(4) = dr.Cells(5).Value ' السعر
                table.Rows(table.Rows.Count - 1)(5) = dr.Cells(6).Value ' الخصم
                table.Rows(table.Rows.Count - 1)(6) = Math.Round((Val(dr.Cells(4).Value) * Val(dr.Cells(5).Value)), 2) ' الاجمالي
                table.Rows(table.Rows.Count - 1)(7) = dr.Cells(11).Value ' المخزن
a:
            Next
            rpt.SetDataSource(table)
            rpt.SetParameterValue("Co_name", My.Settings.Co_name)
            rpt.SetParameterValue("Address", My.Settings.Co_address & " ت : " & My.Settings.Co_tel)
            rpt.SetParameterValue("Bill", txtBill_ID.Text)
            rpt.SetParameterValue("B_Date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
            rpt.SetParameterValue("Customer", cbo_Customer.Text)
            rpt.SetParameterValue("Total", Label18.Text)
            rpt.SetParameterValue("Reduce", Math.Round((Val(Label23.Text) + Val(txtReduce.Text)), 2))
            rpt.SetParameterValue("Total_pur", Label19.Text)
            rpt.SetParameterValue("HR", My.Settings.Co_HR)
            rpt.SetParameterValue("Pay", If(R1.Checked = True, R1.Text, R2.Text))
            rpt.SetParameterValue("Sales_kind", cboPrice.Text)
            rpt.SetParameterValue("Later_acc", Label25.Text)

            If My.Settings.Print = True Then
                rpt.PrintOptions.PrinterName = My.Settings.Printer_Sales
                frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                frmReportViewer.Show()
            Else
                rpt.PrintOptions.PrinterName = My.Settings.Printer_Sales
                rpt.PrintToPrinter(1, False, 0, 0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        MsgBox(Cost, MsgBoxStyle.Information, "رسالة") ' رسالة بها تكلفة الصنف
    End Sub
    Private Sub drg_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellDoubleClick
        drg.CurrentRow.Selected = False ' لآلغاء تحديد الصنف
    End Sub

End Class