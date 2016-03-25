Public Class frmReport_Sales
    Dim Myconn As New Connection
    Dim St As String
    Private Sub Filldrg()
        Try
            Select Case cboSearch.SelectedIndex
                Case 0
                    St = Nothing
                Case 1
                    St = "  and Sales.Users_ID =" & CInt(cboUser.ComboBox.SelectedValue)
                Case 2
                    St = " and Sales.Customer_ID =" & CInt(cboUser.ComboBox.SelectedValue)
                Case 3
                    St = " and Sales.Stock_ID =" & CInt(cboUser.ComboBox.SelectedValue)
            End Select

            drg.Rows.Clear()
            Myconn.ExecQuery("SELECT Sales.ID, Sales.Sales_Date, Sales.Earning, Sales.Sales_time, Sales.Sales_Bill_ID, Sales.items_Cod, Items.Items_Name, Items.Parcode, group.group_Name, Sales.Items_Price, Sales.Items_num, Sales.Total_Price, Sales.Reduce, Sales.Final_Total_Price, Sales.Earning, Sales.Sales_Kind_ID, Sales.Bill_Kind, Sales.Users_ID, Sales.Status, Sales.Customer_ID, Customers.Customer_Name, Users_ID.Employee_Name,Sales.Stock_ID
                                FROM (Users_ID RIGHT JOIN (([group] RIGHT JOIN (Items RIGHT JOIN Sales ON Items.items_Cod = Sales.items_Cod) ON group.Group_cod = Items.Group_cod) RIGHT JOIN Customers ON Sales.Customer_ID = Customers.Customer_ID) ON Users_ID.Employee_ID = Sales.Users_ID) LEFT JOIN Stocks ON Sales.Stock_ID = Stocks.Stock_ID
                                where Sales.Sales_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#" & St & " and Sales.Sales_Bill_ID > 0 order by Sales.Sales_Bill_ID ")

            If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
            Dim V1 As Double = 0
            Dim V2 As Double = 0
            Dim B As Double = 0
            For i As Integer = 0 To Myconn.dt.Rows.Count - 1
                Dim r As DataRow = Myconn.dt.Rows(i)
                drg.Rows.Add()
                drg.Rows(i).Cells(0).Value = i + 1
                drg.Rows(i).Cells(1).Value = r("Sales_Bill_ID")
                drg.Rows(i).Cells(2).Value = r("Customer_Name")
                drg.Rows(i).Cells(3).Value = Format(CDate(r("Sales_Date")), "yyyy/MM/dd")
                drg.Rows(i).Cells(4).Value = r("Sales_time")
                drg.Rows(i).Cells(5).Value = r("Items_Name")
                drg.Rows(i).Cells(6).Value = r("Parcode")
                drg.Rows(i).Cells(7).Value = r("Items_num")
                drg.Rows(i).Cells(8).Value = r("Items_Price")
                drg.Rows(i).Cells(9).Value = r("Reduce")
                drg.Rows(i).Cells(10).Value = r("Final_Total_Price")
                drg.Rows(i).Cells(11).Value = r("Earning")
                drg.Rows(i).Cells(12).Value = r("Employee_Name")
                drg.Rows(i).Cells(13).Value = r("Status")
                drg.Rows(i).Cells(14).Value = r("ID")

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

                If drg.Rows(i).Cells(13).Value = True Then
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.LemonChiffon
                    V1 += CDec(drg.Rows(i).Cells(7).Value) * drg.Rows(i).Cells(8).Value
                    V2 += CDec(drg.Rows(i).Cells(10).Value)
                Else
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    drg.Rows(i).Cells(2).Style.BackColor = Color.Red
                    drg.Rows(i).Cells(5).Style.BackColor = Color.Red
                    B += CDec(drg.Rows(i).Cells(10).Value)
                End If
            Next
            Myconn.DataGridview_MoveLast(drg, 2)
            Label19.Text = V2
            Label18.Text = V1
            Label22.Text = B
            Label23.Text = Math.Round((V1 - V2), 2)
            Label20.Text = "%" & Math.Round(((Val(Val(Label18.Text) - Val(Label19.Text)) / Val(Label18.Text)) * 100), 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
    Private Sub frmReport_Sales_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width

            If F <> 1 Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnSearch.Enabled = r("U_search").ToString
                    btnPrint.Enabled = r("U_print").ToString
                End If
            End If
            'Myconn.Fillcombo("Select * from Users_ID order by Employee_Name", "Users_ID", "Employee_ID", "Employee_Name", Me, cboUser.ComboBox)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        If txt1.Text = "" Then
            MsgBox("أدخل التاريخ", MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة")
            Return
        End If
        If txt2.Text = "" Then
            MsgBox("أدخل التاريخ", MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة")
            Return
        End If
        GroupBox4.Text = "حساب المبيعات خلال الفترة من : " & Format(CDate(txt1.Text), "yyyy/MM/dd") & " إلى " & Format(CDate(txt2.Text), "yyyy/MM/dd") & Space(5)
        Filldrg()

    End Sub
    Private Sub cboSearch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSearch.SelectedIndexChanged
        cboUser.ComboBox.DataSource = Nothing
        Select Case cboSearch.SelectedIndex
            Case 0
                cboUser.Visible = False
            Case 1
                cboUser.Visible = True
                Myconn.Fillcombo("Select * from Users_ID order by Employee_Name", "Users_ID", "Employee_ID", "Employee_Name", Me, cboUser.ComboBox)
            Case 2
                cboUser.Visible = True
                Myconn.Fillcombo("Select * from Customers order by Customer_Name", "Customers", "Customer_ID", "Customer_Name", Me, cboUser.ComboBox)
            Case 3
                cboUser.Visible = True
                Myconn.Fillcombo("Select * from Stocks order by Stock_Name", "Stocks", "Stock_ID", "Stock_Name", Me, cboUser.ComboBox)

        End Select
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim table As New DataTable
        For i As Integer = 1 To 11
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows


            table.Rows.Add()

            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(5).Value
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(6).Value
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(7).Value
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(8).Value
            table.Rows(table.Rows.Count - 1)(5) = " % " & dr.Cells(9).Value
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(10).Value
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(12).Value
            table.Rows(table.Rows.Count - 1)(8) = dr.Cells(3).Value
            table.Rows(table.Rows.Count - 1)(9) = dr.Cells(2).Value
            table.Rows(table.Rows.Count - 1)(10) = dr.Cells(13).Value
        Next
        Dim rpt As New rpt_Sales
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Bill_num", Format(CDate(txt1.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("F_date", Format(CDate(txt2.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("Price", Label18.Text)
        rpt.SetParameterValue("Reduce", Label18.Text - Label19.Text)
        rpt.SetParameterValue("Total", Label19.Text)
        rpt.SetParameterValue("Supplier", Label22.Text)
        rpt.SetParameterValue("Nisba", Label20.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub
End Class