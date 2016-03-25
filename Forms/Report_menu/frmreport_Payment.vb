Public Class frmreport_Payment
    Dim myconn As New Connection
    Dim w, s As Double
    Private Sub Filldrg()
        drg.Rows.Clear()
        myconn.ExecQuery("Select Y.perm_ID,Y.per_time,Y.per_ID,Y.per_date,Y.Amount,Y.Note_per,Y.Status,S.Employee_Name,F.perm_name,(e.Supplier_Name) as Customer_Name,(t.pay_Item_name) as pay_Item_name,Y.pay_Item_ID
                            From (((Safe_payment_per Y Left join Users_ID S on Y.users_ID = S.Employee_ID )
                            Left join Safe_Per F on Y.perm_ID = F.perm_ID)
                            Left join Supplier e on Y.Supplier_ID = e.Supplier_ID)
                            Left join Pay_Items t on Y.pay_Item_ID = t.pay_Item_ID 
                            where Y.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "# and Y.perm_ID > 0 and Y.pay_Item_ID = 1 order by Y.per_date")

        For i As Integer = 0 To myconn.dt.Rows.Count - 1
            Dim r As DataRow = myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("per_time")
            drg.Rows(i).Cells(2).Value = Format(CDate(r("per_date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(3).Value = r("per_ID")
            drg.Rows(i).Cells(4).Value = r("pay_Item_name")
            drg.Rows(i).Cells(5).Value = r("amount")
            drg.Rows(i).Cells(6).Value = clsNumber.nTOword(r("amount"))
            drg.Rows(i).Cells(7).Value = r("Note_per")
            drg.Rows(i).Cells(8).Value = r("Employee_Name")
            drg.Rows(i).Cells(9).Value = r("Status")
            If r("Status") = True Then
                drg.Rows(i).DefaultCellStyle.BackColor = Color.Pink
                w += r("amount")
            Else
                drg.Rows(i).DefaultCellStyle.BackColor = Color.Red
                s += r("amount")
            End If
        Next
        Label18.Text = w
        Label2.Text = s
    End Sub
    Private Sub frmreport_Payment_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width

            If F <> 1 Then
                myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
                Dim r As DataRow = myconn.dt.Rows(0)
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
        If txt1.Text = "" Then
            MsgBox("أدخل التاريخ", MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة")
            Return
        End If
        If txt2.Text = "" Then
            MsgBox("أدخل التاريخ", MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة")
            Return
        End If
        GroupBox2.Text = "المصروفات خلال الفترة من : " & Format(CDate(txt1.Text), "yyyy/MM/dd") & " إلى " & Format(CDate(txt2.Text), "yyyy/MM/dd") & Space(5)
        Filldrg()

    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim table As New DataTable
        For i As Integer = 1 To 10
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows


            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(3).Value
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(7).Value
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(5).Value
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(8).Value

        Next
        Dim rpt As New rpt_Payment
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Bill_num", Format(CDate(txt1.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("F_date", Format(CDate(txt2.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("Price", Label18.Text)
        rpt.SetParameterValue("Reduce", "( " & clsNumber.nTOword(Label18.Text) & " )")

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub

End Class