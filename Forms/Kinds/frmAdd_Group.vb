Public Class frmAdd_Group

    Dim Myconn As New Connection
    Dim fin As String
    Sub Filldrg()
        drg.Rows.Clear()
        Myconn.ExecQuery("SELECT * from [group] order by group_Name")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub

        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("group_Name")
            drg.Rows(i).Cells(2).Value = r("Group_cod")
        Next
        Myconn.DataGridview_MoveLast(drg, 2)
    End Sub
    Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)
        Myconn.Autonumber("Group_cod", "[group]", txtID, Me)
    End Sub
    Sub Save_record()
        With Myconn
            .Parames.Clear()
            .Addparam("@Group_cod", txtID.Text)
            .Addparam("@group_Name", txtName.Text)
            .ExecQuery("insert into  [group] (Group_cod,group_Name) values(@Group_cod,@group_Name)")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
    End Sub
    Private Sub frmAdd_Group_Load(sender As Object, e As EventArgs) Handles Me.Load
        Label5.Left = 0
        Label5.Width = Me.Width
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
            .Addparam("@group_Name", txtName.Text)
            .Addparam("@Group_cod", txtID.Text)
            .ExecQuery("Update  [group]  set group_Name =@group_Name where Group_cod =@Group_cod")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
        drg.CurrentRow.Cells(1).Value = txtName.Text
        MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
    End Sub
    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        With Myconn
            .Addparam("@Group_cod", txtID.Text)
            .ExecQuery("delete from [group] where Group_cod = @Group_cod ")
        End With
        If Myconn.NoErrors(True) = False Then Exit Sub
        drg.Rows.Remove(drg.SelectedRows(0))
        Myconn.ClearAllControls(GroupBox1, True)
    End Sub

    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Myconn.ExecQuery("select * from [group] where Group_cod =" & drg.CurrentRow.Cells(2).Value)
        Dim r As DataRow = Myconn.dt.Rows(0)
        txtID.Text = r("Group_cod").ToString
        txtName.Text = r("group_Name").ToString
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Print_drg()
    End Sub
    Sub Print_drg()
        Try
            Dim rpt As New rpt_Group
            Dim table As New DataTable
            For i As Integer = 1 To 3
                Dim x As String
                x = Format(i, "00")
                table.Columns.Add(x)
            Next

            For Each dr As DataGridViewRow In drg.Rows
                table.Rows.Add()
                table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count ' المسلسل
                table.Rows(table.Rows.Count - 1)(1) = dr.Cells(1).Value ' المجموعة
                table.Rows(table.Rows.Count - 1)(2) = dr.Cells(2).Value ' الباركود

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