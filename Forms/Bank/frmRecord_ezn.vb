Imports System.Globalization
Public Class frmRecord_ezn
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st, st1, st2, st3 As String
    Dim x, y As Integer
    Private Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)
        Myconn.Autonumber("OP_ID", "Bank_Operations", txt_ID, Me)
    End Sub
    Private Sub Filldrg()
        Try
            drg.Rows.Clear()
            Select Case y
                Case 0
                    st = Nothing
                    st1 = Nothing
                    st2 = " where Safe_payment_per.Bank_ID <> NULL"
                    st3 = " where Safe_Recive_per.Bank_ID <> NULL"

                Case 1
                    st = " where Bank_Operations.Op_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
                    st1 = " where Bank_checks.Check_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
                    st2 = " where Safe_payment_per.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "# and Safe_payment_per.Bank_ID <> NULL"
                    st3 = " where Safe_Recive_per.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "# and Safe_Recive_per.Bank_ID <> NULL"

                Case 2
                    st = " where Bank_Operations.Bank_ID =" & CInt(cbo_Bank.SelectedValue)
                    st1 = " where Bank_checks.Bank_ID =" & CInt(cbo_Bank.SelectedValue)
                    st2 = " where Safe_payment_per.Bank_ID =" & CInt(cbo_Bank.SelectedValue)
                    st3 = " where Safe_Recive_per.Bank_ID =" & CInt(cbo_Bank.SelectedValue)

                Case 3
                    st = " where Bank_Operations.Bank_ID =" & CInt(cbo_Bank.SelectedValue) & " and Bank_Operations.Op_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
                    st1 = " where Bank_checks.Bank_ID =" & CInt(cbo_Bank.SelectedValue) & " and Bank_checks.Check_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
                    st2 = " where Safe_payment_per.Bank_ID =" & CInt(cbo_Bank.SelectedValue) & " and Safe_payment_per.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
                    st3 = " where Safe_Recive_per.Bank_ID =" & CInt(cbo_Bank.SelectedValue) & " and Safe_Recive_per.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"

            End Select
            Myconn.ExecQuery("SELECT Bank_Operations.ID, Bank_Operations.Op_date, Bank_Operations.OP_time, Bank_Operations.OP_ID,Bank_Operations.Bank_ID, Bank_Operations.OP_Kind, Bank_Operations.Amount, Bank_Operations.User_ID, Bank_Operations.Notes, Bank.Bank_Name, Users_ID.Employee_Name
                                FROM Users_ID RIGHT JOIN (Bank RIGHT JOIN Bank_Operations ON Bank.Bank_ID = Bank_Operations.Bank_ID) ON Users_ID.Employee_ID = Bank_Operations.User_ID
                                " & st & " order by Bank_Operations.ID 

                                union all

                                SELECT  Bank_checks.ID,Bank_checks.Check_date,'شيك', Bank_checks.Check_ID,  Bank_checks.Bank_ID,Bank_checks.OP_Kind, Bank_checks.Amount,Bank_checks.User_ID,
                                Bank_checks.Notes, Bank.Bank_Name, Users_ID.Employee_Name
                                FROM Users_ID RIGHT JOIN (Bank_checks RIGHT JOIN Bank ON Bank_checks.Bank_ID = Bank.Bank_ID) ON Users_ID.Employee_ID = Bank_checks.User_ID
                                " & st1 & " 

                                union all
                               
                               SELECT  Safe_payment_per.ID, Safe_payment_per.per_date, Safe_payment_per.per_time,Safe_payment_per.per_ID ,Safe_payment_per.Bank_ID ,(0) as OP_Kind,Safe_payment_per.Amount,Safe_payment_per.users_ID, 
                               Safe_payment_per.Note_per, Bank.Bank_Name, Users_ID.Employee_Name
                               FROM Users_ID INNER JOIN (Bank INNER JOIN Safe_payment_per ON Bank.Bank_ID = Safe_payment_per.Bank_ID) ON Users_ID.Employee_ID = Safe_payment_per.users_ID
                               " & st2 & "
                                 
                               union all

                               SELECT Safe_Recive_per.ID, Safe_Recive_per.per_date, Safe_Recive_per.per_time, Safe_Recive_per.perm_ID, Safe_Recive_per.Bank_ID,(1) as OP_Kind , Safe_Recive_per.Amount,
                               Safe_Recive_per.users_ID, Safe_Recive_per.Note_per, Bank.Bank_Name, Users_ID.Employee_Name
                               FROM (Safe_Recive_per LEFT JOIN Bank ON Safe_Recive_per.Bank_ID = Bank.Bank_ID) LEFT JOIN Users_ID ON Safe_Recive_per.users_ID = Users_ID.Employee_ID
                               " & st3)



            If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
            If Myconn.Recodcount = 0 Then
                Select Case y
                    Case 0
                        MsgBox("لا توجد حركات بنكية مسجلة", MsgBoxStyle.Information, "رسالة")
                    Case 1
                        MsgBox("لا  توجد حركات بنكية مسجلة خلال تلك الفترة", MsgBoxStyle.Information, "رسالة")
                    Case 2
                        MsgBox("لا  توجد حركات بنكية مسجلة لهذا البنك", MsgBoxStyle.Information, "رسالة")
                    Case 3
                        MsgBox("لا  توجد حركات بنكية مسجلة لهذا البنك خلال تلك الفترة", MsgBoxStyle.Information, "رسالة")

                End Select
                Return
            End If

            Dim V1 As Double = 0
            Dim V2 As Double = 0
            Dim B As Double = 0
            For i As Integer = 0 To Myconn.dt.Rows.Count - 1
                Dim r As DataRow = Myconn.dt.Rows(i)
                drg.Rows.Add()
                drg.Rows(i).Cells(0).Value = i + 1
                drg.Rows(i).Cells(1).Value = If(r("OP_Kind") = 0, "ايداع", "سحب")
                drg.Rows(i).Cells(2).Value = r("OP_time")
                drg.Rows(i).Cells(3).Value = Format(CDate(r("Op_date").ToString), "yyyy/MM/dd")
                drg.Rows(i).Cells(4).Value = r("OP_ID")
                drg.Rows(i).Cells(5).Value = r("Bank_Name")
                drg.Rows(i).Cells(6).Value = r("Amount")
                drg.Rows(i).Cells(7).Value = r("Notes")
                drg.Rows(i).Cells(8).Value = r("Employee_Name")
                drg.Rows(i).Cells(9).Value = r("ID")

                If r("OP_Kind") = 0 Then
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.LemonChiffon
                    V1 += r("Amount")
                Else
                    drg.Rows(i).DefaultCellStyle.BackColor = Color.Pink
                    V2 += r("Amount")
                End If
            Next
            Myconn.DataGridview_MoveLast(drg, 2)
            B = V1 - V2
            Label11.Text = V1
            Label13.Text = V2
            Label14.Text = B
        Catch ex As Exception
            MsgBox(ex.Message)
            End Try
    End Sub
    Private Sub Binding()
        Try
            Myconn.ExecQuery("Select * from Bank_Operations where ID =" & CInt(drg.CurrentRow.Cells(9).Value))
            If Myconn.Recodcount = 0 Then Return
            Dim r As DataRow = Myconn.dt.Rows(0)
            D_date.Value = r("Op_date")
            txt_ID.Text = r("OP_ID")
            cbo_Bank.SelectedValue = r("Bank_ID")
            cbo_Band.SelectedIndex = r("OP_Kind")
            txtAmount.Text = r("Amount")
            txtNotes.Text = r("Notes")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Save_record()
        With Myconn
            .Parames.Clear()
            .Addparam("@Op_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
            .Addparam("@OP_time", Label12.Text)
            .Addparam("@OP_ID", txt_ID.Text)
            .Addparam("@Bank_ID", cbo_Bank.SelectedValue)
            .Addparam("@OP_Kind", cbo_Band.SelectedIndex)
            .Addparam("@Amount", txtAmount.Text)
            .Addparam("@User_ID", My.Settings.user_ID)
            .Addparam("@Notes", txtNotes.Text)
            .ExecQuery("insert into  [Bank_Operations] (Op_date, OP_time, OP_ID, Bank_ID, OP_Kind, Amount, User_ID, Notes) values(@Op_date,@OP_time,@OP_ID,@Bank_ID,@OP_Kind,@Amount,@User_ID,@Notes)")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
    End Sub
    Private Sub Update_record()
        With Myconn
            .Parames.Clear()
            .Addparam("@Op_date", Format(CDate(D_date.Text), "yyyy/MM/dd"))
            .Addparam("@OP_time", Label12.Text)
            .Addparam("@OP_ID", txt_ID.Text)
            .Addparam("@Bank_ID", cbo_Bank.SelectedValue)
            .Addparam("@OP_Kind", cbo_Band.SelectedIndex)
            .Addparam("@Amount", txtAmount.Text)
            .Addparam("@User_ID", My.Settings.user_ID)
            .Addparam("@Notes", txtNotes.Text)
            .Addparam("@ID", drg.CurrentRow.Cells(9).Value)

            .ExecQuery("Update [Bank_Operations] Set Op_date=@Op_date,OP_time=@OP_time,OP_ID=@OP_ID,Bank_ID=@Bank_ID,OP_Kind=@OP_Kind,Amount=@Amount,User_ID=@User_ID,Notes=@Notes where ID = @ID ")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
    End Sub




    Private Sub frmRecord_ezn_Load(sender As Object, e As EventArgs) Handles Me.Load
        Label5.Left = 0
        Label5.Width = Me.Width
        Try
            If F <> 1 Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " And Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnSave.Enabled = r("U_add").ToString
                    btnSearch.Enabled = r("U_search").ToString
                    btnUpdat.Enabled = r("U_updat").ToString
                    btnNew.Enabled = r("U_new").ToString
                    btnDel.Enabled = r("U_delete").ToString
                    btnPrint.Enabled = r("U_print").ToString
                End If
            End If
        Catch ex As Exception

        End Try

        fin = False
        Myconn.Fillcombo("Select * from [Bank] order by [Bank_name]", "[Bank]", "Bank_ID", "Bank_name", Me, cbo_Bank)
        fin = True
        Timer1.Start()
        y = 0
        Filldrg()

        New_record()
        '-------------------------------------------------------------------------------------------------- النسخة التجريبية
        Myconn.ExecQuery("select * from Bank_Operations")
        If Myconn.Recodcount > 30 Then
            MsgBox("هذه النسخة تجريبية")
            btnSave.Enabled = False
            btnNew.Enabled = False
            Return
        End If
    End Sub
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        New_record()
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is TextBox Then
                If txt.Text = "" And txt.Name <> "txtNotes" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            ElseIf TypeOf txt Is ComboBox Then
                If txt.Text = "" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            End If
        Next
        Save_record()
        y = 2
        Filldrg()
        fin = False
        New_record()
        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
        fin = True
    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        With Myconn
            .Addparam("@ID", drg.CurrentRow.Cells(9).Value)
            .ExecQuery("delete from [Bank_Operations] where ID = @ID ")
        End With
        If Myconn.NoErrors(True) = False Then Exit Sub
        drg.Rows.Remove(drg.SelectedRows(0))
        Myconn.ClearAllControls(GroupBox2, True)
    End Sub

    Private Sub btnUpdat_Click(sender As Object, e As EventArgs) Handles btnUpdat.Click
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is TextBox Then
                If txt.Text = "" And txt.Name <> "txtNotes" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            ElseIf TypeOf txt Is ComboBox Then
                If txt.Text = "" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            End If
        Next
        Update_record()
        y = 0
        Filldrg()
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        If cbo_Bank.SelectedIndex = -1 And String.IsNullOrWhiteSpace(txt1.Text) AndAlso String.IsNullOrWhiteSpace(txt2.Text) Then
            y = 0
        ElseIf String.IsNullOrWhiteSpace(txt1.Text) AndAlso String.IsNullOrWhiteSpace(txt2.Text) And cbo_Bank.SelectedIndex <> -1 Then
            y = 2
        ElseIf Not String.IsNullOrWhiteSpace(txt1.Text) And Not String.IsNullOrWhiteSpace(txt2.Text) And cbo_Bank.SelectedIndex = -1 Then
            y = 1
        ElseIf Not String.IsNullOrWhiteSpace(txt1.Text) And Not String.IsNullOrWhiteSpace(txt2.Text) And cbo_Bank.SelectedIndex <> -1 Then
            y = 3
        End If
            Filldrg()

    End Sub

    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Binding()
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label12.Text = TimeOfDay.ToString("hh:mm:ss tt", CultureInfo.CreateSpecificCulture("ar-eg"))
    End Sub
    Private Sub drg_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellDoubleClick
        drg.CurrentRow.Selected = False
    End Sub

    Private Sub cbo_Bank_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_Bank.SelectedIndexChanged
        If Not fin Then Return
        y = 2
        Filldrg()

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
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value ' الحركة
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value ' الوقت
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' التاريخ
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value 'رقم الاذن 
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value ' البنك
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value 'المبلغ 
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(8).Value ' المستخدم
        Next
        Dim rpt As New rpt_Bank
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        rpt.SetParameterValue("Bill_num", If(cbo_Bank.Text = "", " ", cbo_Bank.Text))
        rpt.SetParameterValue("Price", Label11.Text)
        rpt.SetParameterValue("Reduce", Label13.Text)
        rpt.SetParameterValue("Total", Label14.Text)
        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If
    End Sub
End Class