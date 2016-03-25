Public Class frmReport_Erning
    Dim myconn As New Connection

    Private Sub Filldrg()
        Try
            drg.Rows.Clear()
            myconn.ExecQuery("Select Sum(Final_Total_Price) as Sales_amount,Sum(Earning) as earn,Sales_Date from Sales group by Sales_Date,Status
                            having Sales_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#  and Status = true and min(Sales_Bill_ID) > 0 order by Sales_Date")

            For i As Integer = 0 To myconn.dt.Rows.Count - 1
                Dim r As DataRow = myconn.dt.Rows(i)
                drg.Rows.Add()
                drg.Rows(i).Cells(0).Value = i + 1
                drg.Rows(i).Cells(1).Value = Format(CDate(r("Sales_Date")), "yyyy/MM/dd")
                drg.Rows(i).Cells(2).Value = r("Sales_amount")
                drg.Rows(i).Cells(3).Value = r("earn")
                drg.Rows(i).Cells(4).Value = clsNumber.nTOword(r("earn"))
            Next
            myconn.Sum_drg(drg, 2, Label2, Label3)
            myconn.Sum_drg(drg, 3, Label4, Label6)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub frmReport_Erning_Load(sender As Object, e As EventArgs) Handles Me.Load
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
        GroupBox2.Text = "الأرباح خلال الفترة من : " & Format(CDate(txt1.Text), "yyyy/MM/dd") & " إلى " & Format(CDate(txt2.Text), "yyyy/MM/dd") & Space(5)
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
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value
        Next

        Dim rpt As New rpt_Earning
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Bill_num", Format(CDate(txt1.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("F_date", Format(CDate(txt2.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("Price", Label2.Text)
        rpt.SetParameterValue("Reduce", "( " & clsNumber.nTOword(Label2.Text) & " )")
        rpt.SetParameterValue("Total", Label4.Text)
        rpt.SetParameterValue("Nisba", "( " & clsNumber.nTOword(Label4.Text) & " )")

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub
End Class