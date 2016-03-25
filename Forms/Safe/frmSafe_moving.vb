Imports System.Globalization
Public Class frmSafe_moving
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim St1, St2 As String
    Dim x, y As Integer
    Dim S, W As Double
    Private Sub Fillgrd()
        drg.Rows.Clear()

        Select Case cbo_ezn.SelectedIndex
            Case 0 ' إذن دفع
                Myconn.ExecQuery("Select b.bank_ID,b.bank_name,Y.perm_ID,Y.per_ID,Y.per_date,Y.Amount,Y.Note_per,Y.Status,S.Employee_Name,F.perm_name,(e.Supplier_Name) as Customer_Name,(t.pay_Item_name) as Recive_Item_name
                            From ((((Safe_payment_per Y Left join Users_ID S on Y.users_ID = S.Employee_ID )
                            Left join Safe_Per F on Y.perm_ID = F.perm_ID)
                            Left join Supplier e on Y.Supplier_ID = e.Supplier_ID)
                            Left join Bank b on Y.bank_ID = b.bank_ID)
                            Left join Pay_Items t on Y.pay_Item_ID = t.pay_Item_ID " & St2 & " order by Y.per_date")

            Case 1 ' إذن استلام
                Myconn.ExecQuery("Select b.bank_ID,b.bank_name,R.perm_ID, R.per_ID,R.per_date,R.Amount,R.Note_per,R.Status ,U.Employee_Name,P.perm_name,C.Customer_Name,I.Recive_Item_name
                            From ((((Safe_recive_per R Left join Users_ID U on R.users_ID = U.Employee_ID )
                            Left join Safe_Per P on R.perm_ID = P.perm_ID)
                            Left join Customers C on R.Customer_ID = C.customer_ID)
                            Left join Bank b on R.bank_ID = b.bank_ID)
                            Left join Recive_Items I on R.Recive_Item_ID = I.Recive_Item_ID " & St1 & " order by R.per_date")
            Case 2 '  جميع الأذونات
                Myconn.ExecQuery("Select b.bank_ID,b.bank_name,R.perm_ID, R.per_ID, R.per_date, R.Amount, R.Note_per, R.Status, U.Employee_Name, P.perm_name, C.Customer_Name, I.Recive_Item_name
                            From ((((Safe_recive_per R Left join Users_ID U on R.users_ID = U.Employee_ID )
                            Left join Safe_Per P on R.perm_ID = P.perm_ID)
                            Left join Customers C on R.Customer_ID = C.customer_ID)
                            Left join Bank b on R.bank_ID = b.bank_ID)
                            Left join Recive_Items I on R.Recive_Item_ID = I.Recive_Item_ID " & St1 & "
                            UNION ALL
                            Select b.bank_ID,b.bank_name,Y.perm_ID,Y.per_ID,Y.per_date,Y.Amount,Y.Note_per,Y.Status,S.Employee_Name,F.perm_name,e.Supplier_Name,t.pay_Item_name
                            From ((((Safe_payment_per Y Left join Users_ID S on Y.users_ID = S.Employee_ID )
                            Left join Safe_Per F on Y.perm_ID = F.perm_ID)
                            Left join Supplier e on Y.Supplier_ID = e.Supplier_ID)
                            Left join Bank b on Y.bank_ID = b.bank_ID)
                            Left join Pay_Items t on Y.pay_Item_ID = t.pay_Item_ID " & St2 & " order by R.per_date")
        End Select
        W = 0
        S = 0
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("perm_name")
            drg.Rows(i).Cells(2).Value = r("per_ID")
            drg.Rows(i).Cells(3).Value = r("Recive_Item_name")
            drg.Rows(i).Cells(4).Value = Format(CDate(r("per_date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(5).Value = If(IsDBNull(r("Bank_ID")), If(IsDBNull(r("Customer_Name")), r("Note_per"), r("Customer_Name")), r("Bank_Name"))
            drg.Rows(i).Cells(6).Value = r("amount")
            drg.Rows(i).Cells(7).Value = clsNumber.nTOword(r("amount"))
            drg.Rows(i).Cells(8).Value = r("Employee_Name")
            drg.Rows(i).Cells(9).Value = r("Note_per")
            drg.Rows(i).Cells(10).Value = r("Status")

            If r("perm_ID") = 1 AndAlso r("Status") = True Then
                drg.Rows(i).DefaultCellStyle.BackColor = Color.LemonChiffon
                W += r("amount")
            ElseIf r("perm_ID") = 2 AndAlso r("Status") = True Then
                drg.Rows(i).DefaultCellStyle.BackColor = Color.Pink
                S += r("amount")
            End If
            If r("Status") = False Then
                drg.Rows(i).DefaultCellStyle.BackColor = Color.Red
            End If
        Next

        Label4.Text = W
        Label6.Text = clsNumber.nTOword(Label4.Text)
        Label6.Left = Label4.Left - (Label6.Width + 20)
        Label7.Text = S
        Label8.Text = clsNumber.nTOword(Label7.Text)
        Label8.Left = Label7.Left - (Label8.Width + 20)
        Label9.Text = Math.Round((W - S), 2)
        Label10.Text = clsNumber.nTOword(Label9.Text)
        Label10.Left = Label9.Left - (Label10.Width + 20)
    End Sub
    Private Sub frmSafe_moving_Load(sender As Object, e As EventArgs) Handles Me.Load
        Label5.Left = 0
        Label5.Width = Me.Width

    End Sub
    Private Sub cbo_ezn_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_ezn.SelectedIndexChanged
        Select Case cbo_ezn.SelectedIndex
            Case 0 ' اذن دفع
                St1 = " where R.perm_ID =2"
                St2 = " where Y.perm_ID =2"
            Case 1 ' اذن استلام
                St1 = " where R.perm_ID =1"
                St2 = " where Y.perm_ID =1"
            Case 2 ' كل الاذونات
                St1 = Nothing
                St2 = Nothing
        End Select
        Fillgrd()
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim table As New DataTable
        For i As Integer = 1 To 9
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
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(8).Value
            table.Rows(table.Rows.Count - 1)(8) = dr.Cells(10).Value
        Next
        Dim rpt As New rpt_Safe_move
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Price", Label4.Text)
        rpt.SetParameterValue("reduce", Label7.Text)
        rpt.SetParameterValue("Total", Label9.Text)

        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub

    Private Sub cbo_Search_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_Search.SelectedIndexChanged
        txt1.Visible = False
        txt2.Visible = False
        cbo_band.Visible = False
        cbo_band.ComboBox.DataSource = Nothing
        L1.Visible = False
        L2.Visible = False
        Select Case cbo_Search.SelectedIndex
            Case 0 ' تاريخ
                txt1.Visible = True
            Case 1 ' فترة
                txt2.Visible = True
                txt1.Visible = True
                L1.Visible = True
                L2.Visible = True
            Case 2 ' البند
                Select Case cbo_ezn.SelectedIndex
                    Case 0 ' إذن دفع
                        cbo_band.Visible = True
                        Myconn.Fillcombo("select * from Pay_Items", "Pay_Items", "Pay_Item_ID", "Pay_Item_name", Me, cbo_band.ComboBox)
                    Case 1 ' إذن استلام
                        cbo_band.Visible = True
                        Myconn.Fillcombo("select * from Recive_Items", "Recive_Items", "Recive_Item_ID", "Recive_Item_name", Me, cbo_band.ComboBox)
                    Case 2 '  جميع الأذونات

                End Select

            Case 3 ' بند وتاريخ
                L1.Visible = False
                Select Case cbo_ezn.SelectedIndex
                    Case 0 ' إذن دفع
                        txt1.Visible = True
                        cbo_band.Visible = True
                        Myconn.Fillcombo("select * from Pay_Items", "Pay_Items", "Pay_Item_ID", "Pay_Item_name", Me, cbo_band.ComboBox)
                    Case 1 ' إذن استلام
                        txt1.Visible = True
                        cbo_band.Visible = True
                        Myconn.Fillcombo("select * from Recive_Items", "Recive_Items", "Recive_Item_ID", "Recive_Item_name", Me, cbo_band.ComboBox)
                    Case 2 '  جميع الأذونات

                End Select

            Case 4 ' بند وفترة
                Select Case cbo_ezn.SelectedIndex
                    Case 0 ' إذن دفع
                        txt2.Visible = True
                        txt1.Visible = True
                        L1.Visible = True
                        L2.Visible = True
                        cbo_band.Visible = True
                        Myconn.Fillcombo("select * from Pay_Items", "Pay_Items", "Pay_Item_ID", "Pay_Item_name", Me, cbo_band.ComboBox)
                    Case 1 ' إذن استلام
                        txt2.Visible = True
                        txt1.Visible = True
                        L1.Visible = True
                        L2.Visible = True
                        cbo_band.Visible = True
                        Myconn.Fillcombo("select * from Recive_Items", "Recive_Items", "Recive_Item_ID", "Recive_Item_name", Me, cbo_band.ComboBox)
                    Case 2 '  جميع الأذونات

                End Select
            Case 5 ' عميل
                Select Case cbo_ezn.SelectedIndex
                    Case 0 ' إذن دفع

                    Case 1 ' إذن استلام

                        cbo_band.Visible = True
                        Myconn.Fillcombo("select * from Customers", "Customers", "Customer_ID", "Customer_Name", Me, cbo_band.ComboBox)
                    Case 2 '  جميع الأذونات

                End Select
            Case 6 ' عميل وفترة
                Select Case cbo_ezn.SelectedIndex
                    Case 0 ' إذن دفع

                    Case 1 ' إذن استلام
                        txt2.Visible = True
                        txt1.Visible = True
                        L1.Visible = True
                        L2.Visible = True
                        cbo_band.Visible = True
                        Myconn.Fillcombo("select * from Customers", "Customers", "Customer_ID", "Customer_Name", Me, cbo_band.ComboBox)
                    Case 2 '  جميع الأذونات

                End Select
            Case 7 ' مورد
                Select Case cbo_ezn.SelectedIndex
                    Case 0 ' إذن دفع

                        cbo_band.Visible = True
                        Myconn.Fillcombo("select * from Supplier", "Supplier", "Supplier_ID", "Supplier_Name", Me, cbo_band.ComboBox)
                    Case 1 ' إذن استلام

                    Case 2 '  جميع الأذونات

                End Select
            Case 8 ' مورد وفترة
                Select Case cbo_ezn.SelectedIndex
                    Case 0 ' إذن دفع
                        txt2.Visible = True
                        txt1.Visible = True
                        L1.Visible = True
                        L2.Visible = True
                        cbo_band.Visible = True
                        Myconn.Fillcombo("select * from Supplier", "Supplier", "Supplier_ID", "Supplier_Name", Me, cbo_band.ComboBox)
                    Case 1 ' إذن استلام

                    Case 2 '  جميع الأذونات

                End Select
        End Select
    End Sub
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try

            Select Case cbo_Search.SelectedIndex
                Case 0 ' تاريخ
                    Select Case cbo_ezn.SelectedIndex
                        Case 0 ' إذن دفع
                            St2 = " where Y.per_date = #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "#"
                        Case 1 ' إذن استلام
                            St1 = " where R.per_date = #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "#"
                        Case 2 '  جميع الأذونات
                            St1 = " where R.per_date = #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "#"
                            St2 = " where Y.per_date = #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "#"
                    End Select

                Case 1 ' فترة
                    Select Case cbo_ezn.SelectedIndex
                        Case 0 ' إذن دفع
                            St2 = "where  Y.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"

                        Case 1 ' إذن استلام
                            St1 = "where R.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"

                        Case 2 '  جميع الأذونات
                            St1 = "where R.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"
                            St2 = "where Y.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"
                    End Select

                Case 2 ' البند
                    Select Case cbo_ezn.SelectedIndex
                        Case 0 ' إذن دفع
                            St2 = "where Y.pay_Item_ID =" & CInt(cbo_band.ComboBox.SelectedValue)
                        Case 1 ' إذن استلام
                            St1 = "where R.Recive_Item_ID =" & CInt(cbo_band.ComboBox.SelectedValue)
                        Case 2 '  جميع الأذونات
                            St1 = Nothing
                            St2 = Nothing
                    End Select

                Case 3 ' بند وتاريخ
                    Select Case cbo_ezn.SelectedIndex
                        Case 0 ' إذن دفع
                            St2 = "where Y.pay_Item_ID =" & CInt(cbo_band.ComboBox.SelectedValue) & "and Y.per_date = #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "#"
                        Case 1 ' إذن استلام
                            St1 = "where R.Recive_Item_ID =" & CInt(cbo_band.ComboBox.SelectedValue) & "and R.per_date = #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "#"
                        Case 2 '  جميع الأذونات
                            St1 = Nothing
                            St2 = Nothing
                    End Select

                Case 4 ' بند وفترة
                    Select Case cbo_ezn.SelectedIndex
                        Case 0 ' إذن دفع
                            St2 = "where Y.pay_Item_ID =" & CInt(cbo_band.ComboBox.SelectedValue) & "and Y.per_date  between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"
                        Case 1 ' إذن استلام
                            St1 = "where R.Recive_Item_ID =" & CInt(cbo_band.ComboBox.SelectedValue) & "and R.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"
                        Case 2 '  جميع الأذونات
                            St1 = Nothing
                            St2 = Nothing
                    End Select
                Case 5 ' عميل
                    Select Case cbo_ezn.SelectedIndex
                        Case 0 ' إذن دفع
                            St2 = Nothing

                        Case 1 ' إذن استلام
                            St1 = "where R.Customer_ID =" & CInt(cbo_band.ComboBox.SelectedValue)

                        Case 2 '  جميع الأذونات
                            St1 = Nothing
                            St2 = Nothing
                    End Select
                Case 6 'عميل وفترة
                    Select Case cbo_ezn.SelectedIndex
                        Case 0 ' إذن دفع
                            St2 = Nothing
                        Case 1 ' إذن استلام
                            St1 = "where R.Customer_ID =" & CInt(cbo_band.ComboBox.SelectedValue) & "and R.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"
                        Case 2 '  جميع الأذونات
                            St1 = Nothing
                            St2 = Nothing
                    End Select

                Case 7 ' مورد
                    Select Case cbo_ezn.SelectedIndex
                        Case 0 ' إذن دفع
                            St2 = "where Y.Supplier_ID =" & CInt(cbo_band.ComboBox.SelectedValue)
                        Case 1 ' إذن استلام
                            St1 = Nothing
                        Case 2 '  جميع الأذونات
                            St1 = Nothing
                            St2 = Nothing
                    End Select

                Case 8 ' مورد وفترة
                    Select Case cbo_ezn.SelectedIndex
                        Case 0 ' إذن دفع
                            St1 = "where Y.Supplier_ID =" & CInt(cbo_band.ComboBox.SelectedValue) & "and Y.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"
                        Case 1 ' إذن استلام
                            St1 = Nothing
                        Case 2 '  جميع الأذونات
                            St1 = Nothing
                            St2 = Nothing
                    End Select
            End Select
            Fillgrd()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub drg_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellDoubleClick
        drg.CurrentRow.Selected = False
    End Sub
End Class