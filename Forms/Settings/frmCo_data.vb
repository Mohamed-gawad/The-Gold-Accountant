Public Class frmCo_data
    Private Sub frmCo_data_Load(sender As Object, e As EventArgs) Handles Me.Load
        txtName.Text = My.Settings.Co_name
        txtAddress.Text = My.Settings.Co_address
        txtHR.Text = My.Settings.Co_HR
        txtTel.Text = My.Settings.Co_tel
        txtface.Text = My.Settings.Facebook
        txtMail.Text = My.Settings.Email
        txtSite.Text = My.Settings.Co_Site
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Settings.Co_name = txtName.Text
        My.Settings.Co_address = txtAddress.Text
        My.Settings.Co_HR = txtHR.Text
        My.Settings.Co_tel = txtTel.Text
        My.Settings.Facebook = txtface.Text
        My.Settings.Email = txtMail.Text
        My.Settings.Co_Site = txtSite.Text
        My.Settings.Save()
        MessageBox.Show("تم حفظ الاعدادات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub
End Class