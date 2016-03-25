Public Class frmJob

    Dim Myconn As New Connection

    Private Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)
        Myconn.Autonumber("Job_ID", "Jobs", txtID, Me)
    End Sub
    Private Sub Filldrg()
        Try
            drg.Rows.Clear()
            Myconn.ExecQuery("SELECT * from [Jobs] order by Job_Name")

            If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub

            For i As Integer = 0 To Myconn.dt.Rows.Count - 1
                Dim r As DataRow = Myconn.dt.Rows(i)
                drg.Rows.Add()
                drg.Rows(i).Cells(0).Value = i + 1
                drg.Rows(i).Cells(1).Value = r("Job_Name")
                drg.Rows(i).Cells(2).Value = r("Job_ID")
                drg.Rows(i).Cells(3).Value = r("ID")
            Next
            Myconn.DataGridview_MoveLast(drg, 2)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub Binding()
        Myconn.ExecQuery("SELECT * from [Jobs]  where ID =" & CInt(drg.CurrentRow.Cells(3).Value))
        Dim r As DataRow = Myconn.dt.Rows(0)
        txtID.Text = r("Job_ID").ToString
        txtName.Text = r("Job_Name").ToString

    End Sub
    Private Sub Save_recod()

        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@Job_ID", txtID.Text)
                .Addparam("@Job_Name", txtName.Text)
                .ExecQuery("insert into  [Jobs] (Job_ID,Job_Name)  values(@Job_ID,@Job_Name)")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub Update_record()
        Try
            With Myconn
                .Parames.Clear()
                .Addparam("@Job_Name", txtName.Text)
                .Addparam("@ID", drg.CurrentRow.Cells(3).Value)
                .ExecQuery("Update  [Jobs] set  Job_Name=@Job_Name where ID = @ID")

                If Myconn.NoErrors(True) = False Then Exit Sub
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub frmJob_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width
            If  F <> 1Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnSave.Enabled = r("U_add").ToString
                    btnUpdat.Enabled = r("U_updat").ToString
                    btnNew.Enabled = r("U_new").ToString
                    btnDel.Enabled = r("U_delete").ToString
                    btnPrint.Enabled = r("U_print").ToString
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Filldrg()
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Myconn.ExecQuery("Select * from Jobs where Job_ID =" & CInt(drg.CurrentRow.Cells(2).Value))
        If Myconn.Recodcount > 0 Then MsgBox("هذاالكود مستخدم من قبل") : Return

        If txtID.Text = "" Or txtName.Text = "" Then
            ErrorProvider1.SetError(txtName, "أكمل البيانات")
            MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
            Return

        End If
        Save_recod()
        Filldrg()

        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
        New_record()
    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        Try
            If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
            With Myconn
                .Addparam("@ID", drg.CurrentRow.Cells(3).Value)
                .ExecQuery("delete from [Jobs] where ID = @ID")
            End With
            If Myconn.NoErrors(True) = False Then Exit Sub
            drg.Rows.Remove(drg.SelectedRows(0))
            Myconn.ClearAllControls(GroupBox1, True)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub btnUpdat_Click(sender As Object, e As EventArgs) Handles btnUpdat.Click
        Try
            If txtID.Text = "" Or txtName.Text = "" Then
                ErrorProvider1.SetError(txtName, "أكمل البيانات")
                MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                Return

            End If
            Update_record()
            Filldrg()
            MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Binding()
    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        New_record()

    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            Dim rpt As New rpt_jobs
            Dim table As New DataTable
            For i As Integer = 1 To 3
                Dim x As String
                x = Format(i, "00")
                table.Columns.Add(x)
            Next

            For Each dr As DataGridViewRow In drg.Rows
                table.Rows.Add()
                table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count ' المسلسل
                table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value ' البنك
                table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value ' الكود

            Next
            rpt.SetDataSource(table)
            rpt.SetParameterValue("Co_name", My.Settings.Co_name)
            rpt.SetParameterValue("Address", My.Settings.Co_address & " ت : " & My.Settings.Co_tel)

            rpt.PrintOptions.PrinterName = My.Settings.Printer_Sales
            If My.Settings.Print = True Then
                frmReportViewer.CrystalReportViewer1.ReportSource = rpt
                frmReportViewer.Show()
            Else
                rpt.PrintOptions.PrinterName = My.Settings.Printer_report
                rpt.PrintToPrinter(1, False, 0, 0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "رسالة")
        End Try
    End Sub
End Class