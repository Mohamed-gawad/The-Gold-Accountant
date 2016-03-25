Public Class frmKinds_move
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st, st1 As String

    Private Sub Fillgrd()
        drg.Rows.Clear()

        Myconn.ExecQuery("Select I.Items_Name,I.Parcode,I.items_Cod,Customer_Price, iif(isnull(S.items_Sales),0,S.items_Sales) as items_Sales,iif(isnull(P.items_pur),0,P.items_pur) as items_pur,( iif(isnull(P.items_pur),0,P.items_pur)-iif(isnull(S.items_Sales),0,S.items_Sales)) as rest
                            From (Items I Left join (Select Sum(Items_num) as items_Sales,items_Cod From Sales group by items_Cod,Status having Status = True ) S
                             on I.items_Cod = S.items_Cod )
                            left join (Select Sum(Iteme_Number) as items_pur,items_cod From Purchases group by items_Cod,Status having Status = True ) P
                            on I.items_Cod = P.items_Cod ")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("Items_Name")
            drg.Rows(i).Cells(2).Value = r("Parcode")
            drg.Rows(i).Cells(3).Value = r("Customer_Price")
            drg.Rows(i).Cells(4).Value = r("items_pur")
            drg.Rows(i).Cells(5).Value = r("items_Sales")
            drg.Rows(i).Cells(6).Value = r("rest")
        Next
    End Sub
    Private Sub frmKinds_move_Load(sender As Object, e As EventArgs) Handles Me.Load
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
        Myconn.Fillcombo("Select Items_Name, items_Cod from [items] order by [Items_Name]", "[items]", "items_Cod", "Items_Name", Me, cbo_kind.ComboBox)
        fin = True
        Fillgrd()

    End Sub
    Private Sub cbo_kind_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_kind.SelectedIndexChanged
        If Not fin Then Return
        drg.ClearSelection()
        For W As Integer = 0 To drg.Rows.Count - 1

            If drg.Rows(W).Cells(1).Value.ToString.Equals(cbo_kind.Text, StringComparison.CurrentCultureIgnoreCase) Then
                drg.Rows(W).Cells(2).Selected = True
                drg.CurrentCell = drg.SelectedCells(1)
                Exit For
            End If
        Next

        If cbo_kind.Text = "" Then
            drg.Rows(0).Cells(1).Selected = True
            drg.CurrentCell = drg.SelectedCells(1)
        End If

    End Sub
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Print_Kinds_MOve()
    End Sub
    Private Sub Print_Kinds_MOve() 'حركة الصنف
        Dim table As New DataTable
        For i As Integer = 1 To 7
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value ' الصنف
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value ' الباركود
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' سعر
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value ' المشتريات
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value ' المبيعات
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value ' الرصيد  

        Next
        Dim rpt As New rpt_Kinds_Move
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co_name", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)

        'rpt.SetParameterValue("Bill_num", cbo_kind.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If


    End Sub
    Private Sub txtBarcode_Enter(sender As Object, e As EventArgs) Handles txtBarcode.Enter
        txtBarcode.Text = Nothing
    End Sub
    Private Sub txtBarcode_KeyUp(sender As Object, e As KeyEventArgs) Handles txtBarcode.KeyUp
        If e.KeyCode = Keys.Enter Then
            If Not fin Then Return
            drg.ClearSelection()
            For W As Integer = 0 To drg.Rows.Count - 1

                If drg.Rows(W).Cells(2).Value.ToString.Equals(txtBarcode.Text, StringComparison.CurrentCultureIgnoreCase) Then
                    drg.Rows(W).Cells(2).Selected = True
                    drg.CurrentCell = drg.SelectedCells(1)
                    Exit For
                End If
            Next

            If txtBarcode.Text = "" Then
                drg.Rows(0).Cells(1).Selected = True
                drg.CurrentCell = drg.SelectedCells(1)
            End If
        End If

    End Sub
End Class