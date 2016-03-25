Public Class frmStocks_account
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st, st1 As String

    Private Sub Fillgrd()
        drg.Rows.Clear()

        Myconn.ExecQuery("Select I.Items_Name,I.Parcode,I.items_Cod,I.Customer_Price,I.cost_Price, iif(isnull(S.items_Sales + F.amount_f),0,S.items_Sales + F.amount_f) as items_Sales,iif(isnull(P.items_pur + T.amount_t),0,P.items_pur + T.amount_t) as items_pur,( iif(isnull(P.items_pur + T.amount_t),0,P.items_pur + T.amount_t)-iif(isnull(S.items_Sales + F.amount_f),0,S.items_Sales + F.amount_f)) as rest
                            From (((Items I Left join (Select  iif(isnull(Sum(Items_num)),0,Sum(Items_num)) as items_Sales,items_Cod From Sales group by items_Cod,Status,Stock_ID having Status = True and Stock_ID = " & CInt(cbo_kind.ComboBox.SelectedValue) & " ) S
                             on I.items_Cod = S.items_Cod )
                            left join (Select  iif(isnull(Sum(Iteme_Number)),0,Sum(Iteme_Number)) as items_pur,items_cod From Purchases group by items_Cod,Status,Stock_ID having Status = True and Stock_ID = " & CInt(cbo_kind.ComboBox.SelectedValue) & " ) P
                            on I.items_Cod = P.items_Cod)
                             left join  (Select iif(Isnull(Sum(items_amount)),0,Sum(items_amount)) as amount_F, items_cod From Items_move  group by items_cod,Stock_From having Stock_From = " & CInt(cbo_kind.ComboBox.SelectedValue) & " ) F
                             on  i.items_Cod = F.items_cod)
                             left join  (Select iif(Isnull(Sum(items_amount)),0,Sum(items_amount)) as amount_t, items_cod From Items_move  group by items_cod,Stock_to having Stock_to = " & CInt(cbo_kind.ComboBox.SelectedValue) & " ) T
                            on  i.items_Cod = T.items_cod
                            
                            where iif(isnull(P.items_pur + T.amount_t),0,P.items_pur + T.amount_t) > 0" & st & "")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("Items_Name")
            drg.Rows(i).Cells(2).Value = r("items_Cod")
            drg.Rows(i).Cells(3).Value = r("Parcode")
            drg.Rows(i).Cells(4).Value = If(IsDBNull(r("cost_Price")), 0, r("cost_Price"))
            drg.Rows(i).Cells(5).Value = r("items_pur")
            drg.Rows(i).Cells(6).Value = r("items_Sales")
            drg.Rows(i).Cells(7).Value = r("rest")
            drg.Rows(i).Cells(8).Value = Math.Round((r("rest") * If(IsDBNull(r("cost_Price")), 0, r("cost_Price"))), 2)
        Next
        Myconn.Sum_drg(drg, 8, Label2, Label3)
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        If cboGroup.SelectedIndex = -1 Then
            st = ""
        Else
            st = " and I.group_Cod =" & CInt(cboGroup.ComboBox.SelectedValue)
        End If
        Fillgrd()
    End Sub


    Private Sub frmStocks_account_Load(sender As Object, e As EventArgs) Handles Me.Load
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
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        fin = False
        Myconn.Fillcombo("Select * from [Stocks] order by [Stock_Name]", "[Stocks]", "Stock_ID", "Stock_Name", Me, cbo_kind.ComboBox)
        Myconn.Fillcombo("Select * from [group] order by [group_Name]", "[group]", "Group_cod", "group_Name", Me, cboGroup.ComboBox)
        fin = True

    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim table As New DataTable
        For i As Integer = 1 To 8
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next
        For Each dr As DataGridViewRow In drg.Rows
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count ' لالمسلس
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value 'الصنف
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(3).Value ' الباركود
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(4).Value 'التكلفة
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(5).Value ' المشتريات
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(6).Value ' المبيعات
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(7).Value ' الباقي
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(8).Value ' إجمالي التكلفة
            'table.Rows(table.Rows.Count - 1)(8) = dr.Cells(8).Value '
        Next
        Dim rpt As New rpt_Stocks_account
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Bill_num", If(cbo_kind.Text = "", "جميع المخازن", cbo_kind.Text))
        rpt.SetParameterValue("Price", Label2.Text)
        rpt.SetParameterValue("Supplier", Label3.Text)
        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub
End Class