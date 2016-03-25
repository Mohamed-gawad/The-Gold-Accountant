Public Class frmReport_Purchases
    Dim Myconn As New Connection
    Dim St As String
    Private Sub Filldrg()
        Try
            Select Case cboSearch.SelectedIndex
                Case 0
                    St = Nothing
                Case 1
                    St = " and Purchases.Supplier_ID =" & CInt(cboSupplier.ComboBox.SelectedValue)
            End Select
            drg.Rows.Clear()

            Myconn.ExecQuery("SELECT Purchases.ID, Purchases.Pur_Date, Purchases.Pur_Time, Purchases.Pur_Bill_num, Purchases.Supplier_ID, Purchases.items_Cod, Purchases.Customer_Price, Purchases.Reduce, Purchases.Pur_Price, Purchases.Iteme_Number, Purchases.Total_Price,Purchases.Supplier_ID ,Purchases.Status, Items.Items_Name,Items.Parcode ,group.group_Name, Supplier.Supplier_Name, Users_ID.Employee_Name
                                FROM Users_ID RIGHT JOIN (Supplier RIGHT JOIN ([group] RIGHT JOIN (Items RIGHT JOIN Purchases ON Items.items_Cod = Purchases.items_Cod) ON group.Group_cod = Items.Group_cod) ON Supplier.Supplier_ID = Purchases.Supplier_ID) ON Users_ID.Employee_ID = Purchases.Employee_ID
                               where Purchases.Pur_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "# " & St & " and  Purchases.Pur_Bill_num > 0 order by Purchases.Pur_Bill_num")

            If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
            Dim V1 As Double = 0
            Dim V2 As Double = 0
            Dim B As Double = 0
            For i As Integer = 0 To Myconn.dt.Rows.Count - 1
                Dim r As DataRow = Myconn.dt.Rows(i)
                drg.Rows.Add()
                drg.Rows(i).Cells(0).Value = i + 1
                drg.Rows(i).Cells(1).Value = r("Pur_Bill_num")
                drg.Rows(i).Cells(2).Value = r("Supplier_Name")
                drg.Rows(i).Cells(3).Value = Format(CDate(r("Pur_Date")), "yyyy/MM/dd")
                drg.Rows(i).Cells(4).Value = r("Pur_Time")
                drg.Rows(i).Cells(5).Value = r("Items_Name")
                drg.Rows(i).Cells(6).Value = r("Parcode")
                drg.Rows(i).Cells(7).Value = r("Customer_Price")
                drg.Rows(i).Cells(8).Value = r("Reduce")
                drg.Rows(i).Cells(9).Value = r("Pur_Price")
                drg.Rows(i).Cells(10).Value = r("Iteme_Number")
                drg.Rows(i).Cells(11).Value = r("Total_Price")
                drg.Rows(i).Cells(12).Value = r("Customer_Price") * r("Iteme_Number")
                drg.Rows(i).Cells(13).Value = r("Employee_Name")
                drg.Rows(i).Cells(14).Value = r("Status")
                drg.Rows(i).Cells(15).Value = r("ID")

                If drg.Rows(i).Cells(14).Value = True Then
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.LemonChiffon
                    V1 += CDec(drg.Rows(i).Cells(12).Value)
                    V2 += CDec(drg.Rows(i).Cells(11).Value)
                Else
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    B += CDec(drg.Rows(i).Cells(11).Value)
                End If
            Next
            If My.Settings.Factory_Price = True Then
                drg.Columns(7).Visible = True
                drg.Columns(12).Visible = True
            Else
                drg.Columns(7).Visible = False
                drg.Columns(12).Visible = False
            End If
            Myconn.DataGridview_MoveLast(drg, 2)
            Label18.Text = V1
            Label19.Text = V2
            Label22.Text = B
            Label23.Text = Math.Round((V1 - V2), 2)
            Label20.Text = "%" & Math.Round(((Val(Val(Label18.Text) - Val(Label19.Text)) / Val(Label18.Text)) * 100), 2)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub frmReport_Purchases_Load(sender As Object, e As EventArgs) Handles Me.Load
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
            Myconn.Fillcombo("Select * from Supplier order by Supplier_Name", "Supplier", "Supplier_ID", "Supplier_Name", Me, cboSupplier.ComboBox)
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
        GroupBox4.Text = "حساب المشتريات خلال الفترة من : " & Format(CDate(txt1.Text), "yyyy/MM/dd") & " إلى " & Format(CDate(txt2.Text), "yyyy/MM/dd") & Space(5)
        Filldrg()

    End Sub

    Private Sub cboSearch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSearch.SelectedIndexChanged
        Select Case cboSearch.SelectedIndex
            Case 0
                cboSupplier.Visible = False
            Case 1
                cboSupplier.Visible = True
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
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(10).Value
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(7).Value
            table.Rows(table.Rows.Count - 1)(5) = " % " & dr.Cells(8).Value
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(11).Value
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(13).Value
            table.Rows(table.Rows.Count - 1)(8) = dr.Cells(3).Value
            table.Rows(table.Rows.Count - 1)(9) = dr.Cells(2).Value
            table.Rows(table.Rows.Count - 1)(10) = dr.Cells(14).Value
        Next
        Dim rpt As New rpt_Purchases
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Bill_num", Format(CDate(txt1.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("F_date", Format(CDate(txt2.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("Price", Label18.Text)
        rpt.SetParameterValue("Reduce", Label18.Text - Label19.Text)
        rpt.SetParameterValue("Total", Label19.Text)
        rpt.SetParameterValue("Supplier", " % " & Label22.Text)
        rpt.SetParameterValue("Nisba", " % " & Label20.Text)
        rpt.Section3.ReportObjects.Item("Field2").Border.BackgroundColor = Color.Beige
        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub
End Class