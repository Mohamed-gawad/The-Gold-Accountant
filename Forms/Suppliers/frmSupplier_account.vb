Public Class frmSupplier_account
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st, st1 As String
    Private Sub Filldrg()
        If String.IsNullOrWhiteSpace(txt1.Text) OrElse String.IsNullOrWhiteSpace(txt2.Text) Then
            st = " where Purchases.Supplier_ID =" & CInt(cbo_Customer.ComboBox.SelectedValue) & " and Purchases.Status = True"
        Else
            st = " where Purchases.Supplier_ID =" & CInt(cbo_Customer.ComboBox.SelectedValue) & " and  Purchases.Purchases_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "# and Purchases.Status = True"
        End If

        drg.Rows.Clear()

        Myconn.ExecQuery("Select Items.items_Cod,Items.parcode,Items.Items_Name, Items.cost_Price, Purchases.Pur_Date, Purchases.Pur_Bill_num, Purchases.Customer_Price, Purchases.Iteme_Number, Purchases.Total_Price, Purchases.Reduce, Purchases.Pur_Price, Purchases.Total_Price, Purchases.Employee_ID,Purchases.Status,Purchases.ID ,Supplier.Supplier_Name, Users_ID.Employee_Name,Purchases.Supplier_ID
                            FROM ((Purchases LEFT JOIN Items ON Purchases.items_Cod = Items.items_Cod)
                             LEFT JOIN Supplier ON Purchases.Supplier_ID = Supplier.Supplier_ID)
                             LEFT JOIN Users_ID ON Purchases.Employee_ID = Users_ID.Employee_ID
                             " & st & " order by Purchases.ID ")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        Dim V1 As Double = 0
        Dim V2 As Double = 0

        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = Format(CDate(r("Pur_Date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(2).Value = r("Pur_Bill_num")
            drg.Rows(i).Cells(3).Value = r("Supplier_Name")
            drg.Rows(i).Cells(4).Value = If(IsDBNull(r("items_Cod")), "رصيد أول المدة", r("Items_Name"))
            drg.Rows(i).Cells(5).Value = If(IsDBNull(r("parcode")), 0, r("parcode"))
            drg.Rows(i).Cells(6).Value = If(IsDBNull(r("Customer_Price")), 0, r("Customer_Price"))
            drg.Rows(i).Cells(7).Value = If(IsDBNull(r("Reduce")), 0, r("Reduce"))
            drg.Rows(i).Cells(8).Value = If(IsDBNull(r("Pur_Price")), 0, r("Pur_Price"))
            drg.Rows(i).Cells(9).Value = If(IsDBNull(r("Iteme_Number")), 0, r("Iteme_Number"))
            drg.Rows(i).Cells(10).Value = If(IsDBNull(r("Total_Price")), 0, r("Total_Price"))
            drg.Rows(i).Cells(11).Value = If(IsDBNull(r("Iteme_Number")), 0, r("Iteme_Number")) * If(IsDBNull(r("Customer_Price")), 0, r("Customer_Price"))
            drg.Rows(i).Cells(12).Value = r("Employee_Name")
            drg.Rows(i).Cells(13).Value = r("Status")
            drg.Rows(i).Cells(14).Value = r("ID")
            V1 += r("Total_Price")
        Next
        Label6.Text = V1
        '------------------------------------------------------------------------------------------
        If String.IsNullOrWhiteSpace(txt1.Text) OrElse String.IsNullOrWhiteSpace(txt2.Text) Then
            st = " where  Bank_checks.Supplier_ID =" & CInt(cbo_Customer.ComboBox.SelectedValue)
            st1 = " where  Safe_payment_per.Supplier_ID =" & CInt(cbo_Customer.ComboBox.SelectedValue) & " and Safe_payment_per.Status = True and Safe_payment_per.Pay_Item_ID = 2"
        Else
            st = " where  Bank_checks.Supplier_ID =" & CInt(cbo_Customer.ComboBox.SelectedValue) & " and  Bank_checks.Check_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
            st1 = " where  Safe_payment_per.Supplier_ID =" & CInt(cbo_Customer.ComboBox.SelectedValue) & " and  Safe_payment_per.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "# and Safe_payment_per.Status = True"
        End If

        drg1.Rows.Clear()

        Myconn.ExecQuery("SELECT Safe_payment_per.per_ID, Safe_payment_per.per_date, Safe_payment_per.Amount, Safe_payment_per.Note_per, Supplier.Supplier_Name, Pay_Items.Pay_Item_name, Users_ID.Employee_Name, Safe_payment_per.Supplier_ID
                            FROM Users_ID RIGHT JOIN (Pay_Items RIGHT JOIN (Supplier RIGHT JOIN Safe_payment_per ON Supplier.Supplier_ID = Safe_payment_per.Supplier_ID) ON Pay_Items.Pay_Item_ID = Safe_payment_per.Pay_Item_ID) ON Users_ID.Employee_ID = Safe_payment_per.users_ID
                           " & st1 & " order by Safe_payment_per.ID 

                            Union all

                            SELECT Bank_checks.Check_ID, Bank_checks.Check_date, Bank_checks.Amount, iif(Bank_checks.Notes = '','شيك',Bank_checks.Notes) as Note_per, Supplier.Supplier_Name, Bank.Bank_Name, Users_ID.Employee_Name,Bank_checks.Supplier_ID
                            FROM Bank RIGHT JOIN ((Bank_checks LEFT JOIN Supplier ON Bank_checks.Supplier_ID = Supplier.Supplier_ID) LEFT JOIN Users_ID ON Bank_checks.User_ID = Users_ID.Employee_ID) ON Bank.Bank_ID = Bank_checks.Bank_ID
                          
                            " & st)


        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg1.Rows.Add()
            drg1.Rows(i).Cells(0).Value = i + 1
            drg1.Rows(i).Cells(1).Value = Format(CDate(r("per_date")), "yyyy/MM/dd")
            drg1.Rows(i).Cells(2).Value = r("per_ID")
            drg1.Rows(i).Cells(3).Value = r("supplier_Name")
            drg1.Rows(i).Cells(4).Value = r("Amount")
            drg1.Rows(i).Cells(5).Value = clsNumber.nTOword(r("Amount"))
            drg1.Rows(i).Cells(6).Value = r("pay_Item_name")
            drg1.Rows(i).Cells(7).Value = r("Note_per")
            drg1.Rows(i).Cells(8).Value = r("Employee_Name")
            V2 += r("Amount")
        Next
        Label4.Text = V2
        Label7.Text = Math.Round((V1 - V2), 2)
    End Sub
    Private Sub frmSupplier_account_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width

            If F <> 1 Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " And Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnSearch.Enabled = r("U_search").ToString
                    btnPrint.Enabled = r("U_print").ToString
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        fin = False
        Myconn.Fillcombo("Select Supplier_ID, Supplier_Name from [Supplier] order by [Supplier_Name]", "[Supplier]", "Supplier_ID", "Supplier_Name", Me, cbo_Customer.ComboBox)
        fin = True
    End Sub
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Filldrg()
    End Sub
    Private Sub cbo_Customer_Enter(sender As Object, e As EventArgs) Handles cbo_Customer.Enter
        Myconn.langAR()
    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Select Case cboPrint.SelectedIndex
            Case 0 ' المبيعات
                Print_Pur()
            Case 1 ' المشتريات
                Print_payments()
        End Select
    End Sub

    Private Sub Print_Pur()
        Dim table As New DataTable
        For i As Integer = 1 To 11
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value ' رقم الفاتورة
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(1).Value ' التاريخ
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' المورد
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value ' الصنف
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value ' الباركود
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value ' سعر  
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(9).Value ' العدد
            table.Rows(table.Rows.Count - 1)(8) = Math.Round((dr.Cells(9).Value * dr.Cells(6).Value), 2) ' الاجمالي
            table.Rows(table.Rows.Count - 1)(9) = dr.Cells(7).Value ' الخصم
            table.Rows(table.Rows.Count - 1)(10) = dr.Cells(10).Value ' سعر التكلفة بعد الخصم
        Next
        Dim rpt As New rpt_Supplier_Sales
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)

        rpt.SetParameterValue("Price", Label6.Text)
        rpt.SetParameterValue("Bill_num", cbo_Customer.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub

    Private Sub Print_payments()
        Dim table As New DataTable
        For i As Integer = 1 To 9
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg1.Rows
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            'table.Rows(table.Rows.Count - 1)(1) = dr.Cells(10).Value ' الوقت
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(1).Value 'التاريخ
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(2).Value 'رقم الاذن
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(3).Value ' المورد 
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(4).Value 'المبلغ
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value 'البند
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(8).Value ' المستخدم
            'table.Rows(table.Rows.Count - 1)(8) = dr.Cells(9).Value ' الحالة
        Next
        Dim rpt As New rpt_Supplier_payment
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Price", cbo_Customer.Text)
        rpt.SetParameterValue("Bill_num", Label4.Text)
        rpt.SetParameterValue("Total", clsNumber.nTOword(Label4.Text))
        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub
End Class