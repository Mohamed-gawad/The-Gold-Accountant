Public Class frmBack_Sales
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st As String
    Private Sub Fillgrd()
        drg.Rows.Clear()
        Select Case cboSearch.SelectedIndex
            Case 0 '  مرتجعات صنف
                st = " and Sales.items_Cod = " & CInt(cboSupplier.ComboBox.SelectedValue)
            Case 1 '   صنف خلال فترة
                st = " and Sales.items_Cod = " & CInt(cboSupplier.ComboBox.SelectedValue) & " and Sales.Sales_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
            Case 2 '  مرتجعات خلال فترة
                st = " and Sales.Sales_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
            Case 3 '   مرتجعات مورد
                st = " and Sales.Customer_ID = " & CInt(cboSupplier.ComboBox.SelectedValue)
            Case 4 '   مرتجعات مورد خلال فترة
                st = " and Sales.Customer_ID = " & CInt(cboSupplier.ComboBox.SelectedValue) & " and Sales.Sales_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
            Case 5 '  كل المرتجعات
                st = Nothing
        End Select
        Myconn.ExecQuery("SELECT Sales.Sales_Date, Sales.Sales_time, Sales.Sales_Bill_ID, Sales.items_Cod, Sales.Items_Price, Sales.Items_num, Sales.Total_Price, Sales.Reduce, Sales.Final_Total_Price,
                        Sales.Users_ID, Sales.Stock_ID, Customers.Customer_Name, Items.Items_Name,Sales.Customer_ID,Sales.Items_Price,Users_ID.Employee_Name,Sales.ID
                        FROM ((Sales left join Items on Sales.items_Cod = Items.items_Cod  ) 
                        left join Customers on Customers.Customer_ID = Sales.Customer_ID)
                        left join Users_ID on Sales.Users_ID = Users_ID.Employee_ID where sales.Status = false" & st & " order by Sales.ID")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        Dim V1 As Double = 0
        Dim V2 As Double = 0
        Dim B As Double = 0
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = Format(CDate(r("Sales_Date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(2).Value = r("Sales_Bill_ID")
            drg.Rows(i).Cells(3).Value = r("Customer_Name")
            drg.Rows(i).Cells(4).Value = r("Items_Name")
            drg.Rows(i).Cells(5).Value = r("items_Cod")
            drg.Rows(i).Cells(6).Value = r("Items_Price")
            drg.Rows(i).Cells(7).Value = r("Items_num")
            drg.Rows(i).Cells(8).Value = r("Total_Price")
            drg.Rows(i).Cells(9).Value = r("Reduce")
            drg.Rows(i).Cells(10).Value = r("Final_Total_Price")
            drg.Rows(i).Cells(11).Value = r("Employee_Name")
            drg.Rows(i).Cells(12).Value = r("ID")

            V1 += CDec(drg.Rows(i).Cells(8).Value)
            V2 += CDec(drg.Rows(i).Cells(10).Value)
        Next

        Myconn.DataGridview_MoveLast(drg, 2)
        Label19.Text = V2
        Label18.Text = V1
        Label20.Text = Math.Round(((Val(Val(Label18.Text) - Val(Label19.Text)) / Val(Label18.Text)) * 100), 2)
    End Sub

    Private Sub frmBack_Sales_Load(sender As Object, e As EventArgs) Handles Me.Load
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
    End Sub
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Fillgrd()
    End Sub
    Private Sub cboSearch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSearch.SelectedIndexChanged
        Select Case cboSearch.SelectedIndex
            Case 0 '  مرتجعات صنف
                L1.Visible = False
                L2.Visible = False
                txt1.Visible = False
                txt2.Visible = False
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("select Items_Name,items_Cod from [items] order by [Items_Name]", "[items]", "items_Cod", "Items_Name", Me, cboSupplier.ComboBox)
            Case 1 '   صنف خلال فترة
                L1.Visible = True
                L2.Visible = True
                txt1.Visible = True
                txt2.Visible = True
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("select Items_Name,items_Cod from [items] order by [Items_Name]", "[items]", "items_Cod", "Items_Name", Me, cboSupplier.ComboBox)

            Case 2 '  مرتجعات خلال فترة
                L1.Visible = True
                L2.Visible = True
                txt1.Visible = True
                txt2.Visible = True
                cboSupplier.Visible = False

            Case 3 '   مرتجعات عميل
                L1.Visible = False
                L2.Visible = False
                txt1.Visible = False
                txt2.Visible = False
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("select * from Customers order by Customer_Name", "Customers", "Customer_ID", "Customer_Name", Me, cboSupplier.ComboBox)
            Case 4 '   مرتجعات عميل خلال فترة
                L1.Visible = True
                L2.Visible = True
                txt1.Visible = True
                txt2.Visible = True
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("select * from Customers order by Customer_Name", "Customers", "Customer_ID", "Customer_Name", Me, cboSupplier.ComboBox)

            Case 5 '  كل المرتجعات
                L1.Visible = False
                L2.Visible = False
                txt1.Visible = False
                txt2.Visible = False
                cboSupplier.Visible = False
        End Select

    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Print_bills()

    End Sub
    Private Sub Print_bills()
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
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value ' الكود
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value ' السعر  
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(7).Value 'الكمية
            table.Rows(table.Rows.Count - 1)(8) = dr.Cells(8).Value ' الاجمالي
            table.Rows(table.Rows.Count - 1)(9) = " % " & dr.Cells(9).Value 'الخصم
            table.Rows(table.Rows.Count - 1)(10) = dr.Cells(10).Value ' الاجمالي بعد الخصم


        Next
        Dim rpt As New rpt_Sales_Bills_Back
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Price", Label18.Text)
        rpt.SetParameterValue("Nisba", " % " & Label20.Text)
        rpt.SetParameterValue("Total", Label19.Text)


        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If

    End Sub



End Class