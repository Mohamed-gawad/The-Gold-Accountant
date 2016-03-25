Public Class frmAdd_User

    Dim Myconn As New Connection
    Dim fin As String
    Sub Filldrg()
        drg.Rows.Clear()
        Myconn.ExecQuery("SELECT U.Employee_Name,U.User_Password,U.ID,J.Job_Name from [Users_ID] U Left join Jobs J on U.Job_ID = J.Job_ID order by U.Employee_Name")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub

        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = "*" & r("User_Password") & "*"
            drg.Rows(i).Cells(1).Value = r("Job_Name")
            drg.Rows(i).Cells(2).Value = r("Employee_Name")
            drg.Rows(i).Cells(3).Value = i + 1
            drg.Rows(i).Cells(4).Value = r("ID")
            drg.Rows(i).Cells(5).Value = r("User_Password")
        Next
        Myconn.DataGridview_MoveLast(drg, 2)
    End Sub
    Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)
        'Myconn.Autonumber("Pass_ID", "[Users_ID]", txtID, Me)
    End Sub
    Sub Save_record()
        Myconn.ExecQuery("Select * from Users_ID where Employee_ID =" & CInt(cbo_Employee.SelectedValue))
        If Myconn.Recodcount > 0 Then MsgBox("هذا المستخدم تمت إضافته من قبل") : Return
        With Myconn
            .Parames.Clear()
            .Addparam("@Employee_ID", cbo_Employee.SelectedValue)
            .Addparam("@Employee_Name", cbo_Employee.Text)
            .Addparam("@User_Password", txtPass.Text)
            .Addparam("@job_ID", cbo_Job.SelectedValue)
            .ExecQuery("insert into  [Users_ID] (Employee_ID,Employee_Name,User_Password,job_ID) values(@Employee_ID,@Employee_Name,@User_Password,@job_ID)")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
    End Sub

    Private Sub frmAdd_User_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width
            If F = 0 Then GoTo W
            If F <> 1 Then
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
            New_record()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
W:
        Myconn.Fillcombo("Select * from Employees order by Employee_name", "Employees", "Employee_ID", "Employee_name", Me, cbo_Employee)
        Myconn.Fillcombo("Select * from Jobs order by job_name", "Jobe", "Job_Id", "Job_Name", Me, cbo_Job)
        Filldrg()
    End Sub
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        New_record()
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is TextBox Then
                If txt.Text = "" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            End If
        Next
        Save_record()
        Filldrg()
        New_record()
    End Sub
    Private Sub btnUpdat_Click(sender As Object, e As EventArgs) Handles btnUpdat.Click
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is TextBox Then
                If txt.Text = "" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            End If
        Next

        With Myconn
            .Parames.Clear()
            .Addparam("@Employee_ID", cbo_Employee.SelectedValue)
            .Addparam("@Employee_Name", cbo_Employee.Text)
            .Addparam("@User_Password", txtPass.Text)
            .Addparam("@job_ID", cbo_Job.SelectedValue)
            .Addparam("@ID", drg.CurrentRow.Cells(4).Value)
            .ExecQuery("Update [Users_ID]  set Employee_ID=@Employee_ID,Employee_Name=@Employee_Name,User_Password=@User_Password,job_ID=@job_ID where ID =@ID")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
        Filldrg()

        MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
    End Sub
    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        With Myconn
            .Addparam("@ID", drg.CurrentRow.Cells(4).Value)
            .ExecQuery("delete from [Users_ID] where ID = @ID ")
        End With
        If Myconn.NoErrors(True) = False Then Exit Sub
        drg.Rows.Remove(drg.SelectedRows(0))
        Myconn.ClearAllControls(GroupBox1, True)
    End Sub

    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Myconn.ExecQuery("select * from [Users_ID] where ID =" & drg.CurrentRow.Cells(4).Value)
        Dim r As DataRow = Myconn.dt.Rows(0)
        cbo_Employee.SelectedValue = r("Employee_ID").ToString
        txtPass.Text = r("User_Password")
        cbo_Job.SelectedValue = r("job_ID")
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            Dim rpt As New rpt_Users
            Dim table As New DataTable
            For i As Integer = 1 To 4
                Dim x As String
                x = Format(i, "00")
                table.Columns.Add(x)
            Next

            For Each dr As DataGridViewRow In drg.Rows
                table.Rows.Add()
                table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count ' المسلسل
                table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value ' المستخدم
                table.Rows(table.Rows.Count - 1)(2) = dr.Cells(1).Value ' الوظيفة
                table.Rows(table.Rows.Count - 1)(3) = dr.Cells(5).Value ' الباسورد
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