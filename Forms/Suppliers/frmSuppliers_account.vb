Public Class frmSuppliers_account
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st, st1 As String
    Private Sub Fillgrd()
        drg.Rows.Clear()

        Myconn.ExecQuery("Select C.Supplier_Name,C.Supplier_ID,iif(isnull(S.Total),0,S.Total) as Total,(iif(isnull(P.Amount2),0,P.Amount2) + iif(isnull(B.cheek),0,B.cheek)) as Amount,(iif(isnull(S.Total),0,S.Total) - iif(isnull(P.Amount2),0,P.Amount2) + iif(isnull(B.cheek),0,B.cheek)) as rest
                             From ((Supplier C Left join (Select Sum(Total_Price) as Total, Supplier_ID From Purchases group by Supplier_ID,Status having Status = True ) S
                             on C.Supplier_ID = S.Supplier_ID )
                            left join (Select Sum(Amount) as Amount2,Supplier_ID From Safe_payment_per group by Supplier_ID,Status having Status = True ) P
                            on C.Supplier_ID = P.Supplier_ID )
                            left join (Select Sum(Amount) as cheek,Supplier_ID from Bank_checks group by Supplier_ID ) B
                            on C.Supplier_ID= B.Supplier_ID")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("Supplier_Name")
            drg.Rows(i).Cells(2).Value = r("Supplier_ID")
            drg.Rows(i).Cells(3).Value = r("Total")
            drg.Rows(i).Cells(4).Value = r("Amount")
            drg.Rows(i).Cells(5).Value = r("rest")
            If r("rest") > 0 Then
                st = "دائن"
            ElseIf r("rest") = 0 Then
                st = "خالص"
            ElseIf r("rest") < 0 Then
                st = "مدين"
            End If

            drg.Rows(i).Cells(6).Value = st
        Next
        Myconn.Sum_drg(drg, 5, Label4, Label6)
    End Sub
    Private Sub frmSuppliers_account_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width

            If F <> 1 Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnPrint.Enabled = r("U_print").ToString
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        fin = False
        Myconn.Fillcombo("Select Supplier_ID, Supplier_Name from [Supplier] order by [Supplier_Name]", "[Suppliers]", "Supplier_ID", "Supplier_Name", Me, cbo_Customer.ComboBox)
        fin = True
        Fillgrd()

    End Sub
    Private Sub cbo_Customer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_Customer.SelectedIndexChanged
        If Not fin Then Return
        drg.ClearSelection()
        For W As Integer = 0 To drg.Rows.Count - 1

            If drg.Rows(W).Cells(1).Value.ToString.Equals(cbo_Customer.Text, StringComparison.CurrentCultureIgnoreCase) Then
                drg.Rows(W).Cells(2).Selected = True
                drg.CurrentCell = drg.SelectedCells(1)
                Exit For
            End If
        Next

        If cbo_Customer.Text = "" Then
            drg.Rows(0).Cells(1).Selected = True
            drg.CurrentCell = drg.SelectedCells(1)
        End If

    End Sub
    Private Sub cbo_Customer_Enter(sender As Object, e As EventArgs) Handles cbo_Customer.Enter
        Myconn.langAR()
    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim table As New DataTable
        For i As Integer = 1 To 7
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value ' المورد
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value ' الكود
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' المبيعات
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value ' المدفوعات
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value 'الباقي
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value ' الحالة
        Next
        Dim rpt As New rpt_Suppliers_accounts
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co_name", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Total", Label4.Text)
        rpt.SetParameterValue("Total_pur", Label6.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub



End Class