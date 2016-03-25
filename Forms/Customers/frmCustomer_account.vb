Public Class frmCustomer_account
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st, st1 As String
    Private Sub Filldrg()
        If String.IsNullOrWhiteSpace(txt1.Text) OrElse String.IsNullOrWhiteSpace(txt2.Text) Then
            st = " where Sales.Customer_ID =" & CInt(cbo_Customer.ComboBox.SelectedValue) & " and Sales.Status = True"
        Else
            st = " where Sales.Customer_ID =" & CInt(cbo_Customer.ComboBox.SelectedValue) & " and  Sales.Sales_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "# and Sales.Status = True"
        End If

        drg.Rows.Clear()
        Myconn.ExecQuery("Select Items.items_Cod, Items.Items_Name,Items.Parcode ,Items.cost_Price, Sales.Sales_Date, Sales.Sales_Bill_ID, Sales.Items_Price, Sales.Items_num, Sales.Total_Price, Sales.Reduce, Sales.Final_Total_Price, Sales.Final_Total_Price, Sales.Users_ID,Sales.Status,sales.ID ,Sales.Earning, Customers.Customer_Name, Users_ID.Employee_Name,Sales.Customer_ID
                            FROM ((Sales LEFT JOIN Items ON Sales.items_Cod = Items.items_Cod) LEFT JOIN Customers ON Sales.Customer_ID = Customers.Customer_ID) LEFT JOIN Users_ID ON Sales.Users_ID = Users_ID.Employee_ID
                             " & st & " order by Sales.ID ")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        Dim V1 As Double = 0
        Dim V2 As Double = 0

        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = Format(CDate(r("Sales_Date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(2).Value = r("Sales_Bill_ID")
            drg.Rows(i).Cells(3).Value = r("Customer_Name")
            drg.Rows(i).Cells(4).Value = If(IsDBNull(r("items_Cod")), "رصيد أول المدة", r("Items_Name"))
            drg.Rows(i).Cells(5).Value = r("Parcode")
            drg.Rows(i).Cells(6).Value = r("Items_num")
            drg.Rows(i).Cells(7).Value = r("Items_Price")
            drg.Rows(i).Cells(8).Value = r("Reduce")
            drg.Rows(i).Cells(9).Value = r("Final_Total_Price")
            drg.Rows(i).Cells(10).Value = r("Employee_Name")
            drg.Rows(i).Cells(11).Value = r("Status")
            drg.Rows(i).Cells(12).Value = r("ID")
            V1 += r("Final_Total_Price")
        Next
        Label6.Text = V1
        '------------------------------------------------------------------------------------------
        If String.IsNullOrWhiteSpace(txt1.Text) OrElse String.IsNullOrWhiteSpace(txt2.Text) Then
            st1 = " where  Safe_Recive_per.Customer_ID =" & CInt(cbo_Customer.ComboBox.SelectedValue)
        Else
            st1 = " where  Safe_Recive_per.Customer_ID =" & CInt(cbo_Customer.ComboBox.SelectedValue) & " and  Safe_Recive_per.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
        End If

        drg1.Rows.Clear()
        Myconn.ExecQuery("SELECT Safe_Recive_per.per_ID,Safe_Recive_per.Status,Safe_Recive_per.per_time ,Safe_Recive_per.per_date, Safe_Recive_per.Amount, Safe_Recive_per.Note_per, Customers.Customer_Name, Recive_Items.Recive_Item_name, Users_ID.Employee_Name, Safe_Recive_per.Customer_ID
                            FROM Users_ID RIGHT JOIN (Recive_Items RIGHT JOIN (Customers RIGHT JOIN Safe_Recive_per ON Customers.Customer_ID = Safe_Recive_per.Customer_ID) ON Recive_Items.Recive_Item_ID = Safe_Recive_per.Recive_Item_ID) ON Users_ID.Employee_ID = Safe_Recive_per.users_ID
                            " & st1 & " order by Safe_Recive_per.ID ")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg1.Rows.Add()
            drg1.Rows(i).Cells(0).Value = i + 1
            drg1.Rows(i).Cells(1).Value = Format(CDate(r("per_date")), "yyyy/MM/dd")
            drg1.Rows(i).Cells(2).Value = r("per_ID")
            drg1.Rows(i).Cells(3).Value = r("Customer_Name")
            drg1.Rows(i).Cells(4).Value = r("Amount")
            drg1.Rows(i).Cells(5).Value = clsNumber.nTOword(r("Amount"))
            drg1.Rows(i).Cells(6).Value = r("Recive_Item_name")
            drg1.Rows(i).Cells(7).Value = r("Note_per")
            drg1.Rows(i).Cells(8).Value = r("Employee_Name")
            drg1.Rows(i).Cells(9).Value = r("Status")
            drg1.Rows(i).Cells(10).Value = r("per_time")
            If drg1.Rows(i).Cells(9).Value = True Then
                drg1.Rows(i).DefaultCellStyle.BackColor = Color.LemonChiffon
                V2 += r("Amount")
            Else
                drg1.Rows(i).DefaultCellStyle.BackColor = Color.Red
            End If
        Next
        Label4.Text = V2
        Label7.Text = V1 - V2
    End Sub
    Private Sub frmCustomer_account_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width

            If  F <> 1Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
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
        Myconn.Fillcombo("Select Customer_ID, Customer_Name from [Customers] order by [Customer_Name]", "[Customers]", "Customer_ID", "Customer_Name", Me, cbo_Customer.ComboBox)
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
            Case 0 ' المشتريات
                Print_Pur()
            Case 1 ' المدفوعات
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
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' العميل
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value ' الصنف
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value ' الباركود
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(7).Value ' سعر  
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(6).Value ' العدد
            table.Rows(table.Rows.Count - 1)(8) = Math.Round((dr.Cells(7).Value * dr.Cells(6).Value), 2) ' الاجمالي
            table.Rows(table.Rows.Count - 1)(9) = dr.Cells(8).Value ' الخصم
            table.Rows(table.Rows.Count - 1)(10) = dr.Cells(9).Value ' سعر التكلفة بعد الخصم
        Next
        Dim rpt As New rpt_Customer_Sales
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
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(10).Value ' الوقت
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(1).Value 'التاريخ
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(2).Value 'رقم الاذن
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(3).Value ' العميل 
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(4).Value 'المبلغ
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value 'البند
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(8).Value ' المستخدم
            table.Rows(table.Rows.Count - 1)(8) = dr.Cells(9).Value ' الحالة
        Next
        Dim rpt As New rpt_Customer_payment
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