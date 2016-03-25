Public Class frmAll_Sales_Bills
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st As String

    Private Sub fillgrd()
        drg.Rows.Clear()

        Select Case cboSearch.SelectedIndex
            Case 0 '  كل الفواتير
                st = Nothing
            Case 1 ' فواتير خلال فترة
                st = " and S.Sales_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
            Case 2 ' فواتير مورد
                st = " and S.Customer_ID = " & CInt(cboSupplier.ComboBox.SelectedValue)
            Case 3 ' فواتير مورد خلال فترة
                st = " and S.Sales_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#" & " and S.Customer_ID = " & CInt(cboSupplier.ComboBox.SelectedValue)
        End Select
        Myconn.ExecQuery("SELECT C.Customer_Name,S.Customer_ID,S.Sales_Date,S.Sales_Bill_ID,count(S.items_Cod) as items_num,sum(S.Total_Price) as Total,
                            iif(isnull(B.back),0,B.back) as back,iif(isnull(B.back_cost),0,B.back_cost) as back_cost,S.Status,sum(s.Final_Total_Price) as Final_Total_Price
                            from (Sales S left join (Select count(items_Cod) as back ,sum(Final_Total_Price) as back_cost,Sales_Bill_ID from Sales group by Sales_Bill_ID , Status having Status = false ) B
                            on S.Sales_Bill_ID = b.Sales_Bill_ID)
                            left join Customers C on S.Customer_ID = C.Customer_ID
                            group by C.Customer_Name,S.Customer_ID,S.Sales_Date,S.Sales_Bill_ID,B.back,B.back_cost,S.status having S.Status = true " & st & " and S.Sales_Bill_ID > 0 order by S.Sales_Bill_ID")


        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        Dim V1 As Double = 0
        Dim V2 As Double = 0
        Dim B As Double = 0
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("Sales_Bill_ID")
            drg.Rows(i).Cells(2).Value = Format(CDate(r("Sales_Date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(3).Value = r("Customer_Name")
            drg.Rows(i).Cells(4).Value = r("Customer_ID")
            drg.Rows(i).Cells(5).Value = r("items_num")
            drg.Rows(i).Cells(6).Value = r("back")
            drg.Rows(i).Cells(7).Value = r("back_cost")
            drg.Rows(i).Cells(8).Value = r("Total")
            drg.Rows(i).Cells(9).Value = Math.Round(((((r("Total") - r("Final_Total_Price")) / r("Total"))) * 100), 2)
            drg.Rows(i).Cells(10).Value = r("Final_Total_Price")

            V1 += r("Final_Total_Price")
            V2 += r("Total")
            B += r("back_cost")
        Next
        Myconn.DataGridview_MoveLast(drg, 2)
        Label18.Text = V2
        Label19.Text = V1
        Label22.Text = B
        Label20.Text = Math.Round(((Val(Val(Label18.Text) - Val(Label19.Text)) / Val(Label18.Text)) * 100), 2)
    End Sub

    Private Sub frmAll_Sales_Bills_Load(sender As Object, e As EventArgs) Handles Me.Load

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
        fillgrd()
    End Sub

    Private Sub cboSearch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSearch.SelectedIndexChanged
        Select Case cboSearch.SelectedIndex
            Case 0 ' جميع الفواتير
                cboSupplier.Visible = False
                L1.Visible = False
                L2.Visible = False
                txt1.Visible = False
                txt2.Visible = False

            Case 1 ' فواتير خلال فترة
                cboSupplier.Visible = False
                L1.Visible = True
                L2.Visible = True
                txt1.Visible = True
                txt2.Visible = True
            Case 2 ' فواتير عميل
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("Select * from customers order by customer_name", "customers", "customer_ID", "customer_name", Me, cboSupplier.ComboBox)
                L1.Visible = False
                L2.Visible = False
                txt1.Visible = False
                txt2.Visible = False
            Case 3 ' فواتير عميل خلال فترة
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("Select * from customers order by customer_name", "customers", "customer_ID", "customer_name", Me, cboSupplier.ComboBox)
                L1.Visible = True
                L2.Visible = True
                txt1.Visible = True
                txt2.Visible = True
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
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value '
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value '
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' 
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value ' 
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value ' 
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value '  
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(7).Value '
            table.Rows(table.Rows.Count - 1)(8) = dr.Cells(8).Value '
            table.Rows(table.Rows.Count - 1)(9) = " % " & dr.Cells(9).Value '
            table.Rows(table.Rows.Count - 1)(10) = dr.Cells(10).Value '


        Next
        Dim rpt As New rpt_Sales_Bills
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        'rpt.SetParameterValue("Bill_num", Format(CDate(txt1.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("Price", Label18.Text)
        rpt.SetParameterValue("reduce", Label18.Text - Label19.Text)
        rpt.SetParameterValue("Nisba", Label20.Text)
        rpt.SetParameterValue("Total", Label19.Text)
        rpt.SetParameterValue("Supplier", Label22.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If

    End Sub
End Class