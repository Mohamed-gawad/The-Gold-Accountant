Public Class frmAdd_Employee

    Dim Myconn As New Connection
    Dim fin As String
    Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)
        Myconn.Autonumber("Employee_ID", "[Employees]", txtID, Me)
    End Sub
    Sub Filldrg()
        drg.Rows.Clear()
        Myconn.ExecQuery("SELECT E.Employee_ID,E.Employee_Name,E.Employee_NID,E.Employee_address,E.Employee_tel1,E.Employee_tel2,
                            E.Employee_salary,E.Employee_hours,E.ID,J.Job_Name from [Employees] E left join Jobs J on E.Job_ID = j.job_ID order by E.Employee_ID")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub

        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("Employee_Name")
            drg.Rows(i).Cells(2).Value = r("Employee_ID")
            drg.Rows(i).Cells(3).Value = r("Job_Name")
            drg.Rows(i).Cells(4).Value = r("Employee_NID")
            drg.Rows(i).Cells(5).Value = r("Employee_address")
            drg.Rows(i).Cells(6).Value = r("Employee_tel1")
            drg.Rows(i).Cells(7).Value = r("Employee_salary")
            drg.Rows(i).Cells(8).Value = r("Employee_hours")
            drg.Rows(i).Cells(9).Value = r("ID")
        Next
        Myconn.DataGridview_MoveLast(drg, 2)
    End Sub
    Sub Save_record()
        With Myconn
            .Parames.Clear()
            .Addparam("@Job_ID", cboJobs.SelectedValue)
            .Addparam("@Employee_Name", txtName.Text)
            .Addparam("@Employee_ID", CInt(txtID.Text))
            .Addparam("@Employee_NID", txt_NID.Text)
            .Addparam("@Employee_address", txtAddress.Text)
            .Addparam("@Employee_tel1", txtTel1.Text)
            .Addparam("@Employee_tel2", txtTel2.Text)
            .Addparam("@Employee_salary", CDbl(txtSalary.Text))
            .Addparam("@Employee_hours", CInt(txtHours.Text))

            .ExecQuery("insert into  [Employees] (Job_ID,Employee_Name,Employee_ID,Employee_NID,Employee_address,Employee_tel1,Employee_tel2,Employee_salary,Employee_hours) 
                                           values(@Job_ID,@Employee_Name,@Employee_ID,@Employee_NID,@Employee_address,@Employee_tel1,@Employee_tel2,@Employee_salary,@Employee_hours)")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
    End Sub
    Private Sub frmAdd_Employee_Load(sender As Object, e As EventArgs) Handles Me.Load
        Label5.Left = 0
        Label5.Width = Me.Width
        Myconn.Fillcombo("Select * from jobs", "jobs", "job_ID", "job_Name", Me, cboJobs)
        Filldrg()
        New_record()
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
        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

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
            .Addparam("@Job_ID", cboJobs.SelectedValue)
            .Addparam("@Employee_Name", txtName.Text)
            .Addparam("@Employee_ID", txtID.Text)
            .Addparam("@Employee_NID", txt_NID.Text)
            .Addparam("@Employee_address", txtAddress.Text)
            .Addparam("@Employee_tel1", txtTel1.Text)
            .Addparam("@Employee_tel2", txtTel2.Text)
            .Addparam("@Employee_salary", txtSalary.Text)
            .Addparam("@Employee_hours", txtHours.Text)
            .Addparam("@ID", drg.CurrentRow.Cells(9).Value)
            .ExecQuery("Update  [Employees]  set Job_ID=@Job_ID,Employee_Name=@Employee_Name,Employee_ID=@Employee_ID,Employee_NID=@Employee_NID,Employee_address=@Employee_address,Employee_tel1=@Employee_tel1,Employee_tel2=@Employee_tel2,Employee_salary=@Employee_salary,Employee_hours=@Employee_hours where ID=@ID")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
        Filldrg()

        MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
    End Sub
    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        With Myconn
            .Addparam("@ID", drg.CurrentRow.Cells(9).Value)
            .ExecQuery("delete from [Employees] where ID = @ID")
        End With
        If Myconn.NoErrors(True) = False Then Exit Sub
        drg.Rows.Remove(drg.SelectedRows(0))
        Myconn.ClearAllControls(GroupBox1, True)
    End Sub
    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Myconn.ExecQuery("select * from [Employees] where ID =" & drg.CurrentRow.Cells(9).Value)
        Dim r As DataRow = Myconn.dt.Rows(0)
        txtID.Text = r("Employee_ID").ToString
        txtName.Text = r("Employee_Name").ToString
        txt_NID.Text = r("Employee_NID").ToString
        txtAddress.Text = r("Employee_address").ToString
        txtTel1.Text = r("Employee_tel1").ToString
        txtTel2.Text = r("Employee_tel2").ToString
        txtSalary.Text = r("Employee_salary").ToString
        txtHours.Text = r("Employee_hours").ToString
        cboJobs.SelectedValue = r("Job_ID")
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            Dim rpt As New rpt_Employees
            Dim table As New DataTable
            For i As Integer = 1 To 6
                Dim x As String
                x = Format(i, "00")
                table.Columns.Add(x)
            Next

            For Each dr As DataGridViewRow In drg.Rows
                table.Rows.Add()
                table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count ' المسلسل
                table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value ' العميل
                table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value ' الكود
                table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' الوظيفة
                table.Rows(table.Rows.Count - 1)(4) = dr.Cells(5).Value ' العنوان
                table.Rows(table.Rows.Count - 1)(5) = dr.Cells(6).Value ' المحمول

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