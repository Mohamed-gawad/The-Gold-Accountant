Imports System.Drawing.Printing
Public Class frmBarcode_Setting
    Private Sub frmBarcode_Setting_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim pkInstalledPrinters As String

        ' Find all printers installed
        For Each pkInstalledPrinters In
            PrinterSettings.InstalledPrinters
            cboPrinter.Items.Add(pkInstalledPrinters)

        Next pkInstalledPrinters

        cboPrinter.Text = My.Settings.Printer_barcode
        cboLabel_Size.Text = My.Settings.B_Size
        cbo_L1.SelectedIndex = My.Settings.B_L1
        cbo_L3.SelectedIndex = My.Settings.B_L3
        cbo_L2.SelectedIndex = 0
        cbo_L4.SelectedIndex = My.Settings.B_L4
        cbo_L5.SelectedIndex = My.Settings.B_L5

        cboLabel_Size.SelectedIndex = My.Settings.B_Size
        txtNum.Text = My.Settings.label_number
        CheckBox1.Checked = My.Settings.B_L1_V
        CheckBox3.Checked = My.Settings.B_L3_V
        CheckBox4.Checked = My.Settings.B_L4_V
        CheckBox5.Checked = My.Settings.B_L5_v
        CheckBox6.Checked = My.Settings.Barcode_Previo

        txtTop.Text = My.Settings.M_Top
        txtButtom.Text = My.Settings.M_Butom
        txtRight.Text = My.Settings.M_Right
        txtLeft.Text = My.Settings.M_Left

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        My.Settings.Printer_barcode = cboPrinter.Text
        My.Settings.B_Size = cboLabel_Size.SelectedIndex
        My.Settings.B_L1 = cbo_L1.SelectedIndex
        My.Settings.B_L3 = cbo_L3.SelectedIndex
        My.Settings.B_L4 = cbo_L4.SelectedIndex
        My.Settings.B_L5 = cbo_L5.SelectedIndex
        My.Settings.B_L1_V = CheckBox1.Checked
        My.Settings.label_number = txtNum.Text
        My.Settings.B_L3_V = CheckBox3.Checked
        My.Settings.B_L4_V = CheckBox4.Checked
        My.Settings.B_L5_v = CheckBox5.Checked
        My.Settings.Barcode_Previo = CheckBox6.Checked

        My.Settings.M_Top = txtTop.Text
        My.Settings.M_Butom = txtButtom.Text
        My.Settings.M_Right = txtRight.Text
        My.Settings.M_Left = txtLeft.Text

        My.Settings.Save()

        MessageBox.Show("تم حفظ الاعدادات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)


    End Sub

    Private Sub cboLabel_Size_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLabel_Size.SelectedIndexChanged
        Select Case cboLabel_Size.SelectedIndex
            Case 0
                txtNum.Enabled = True
            Case 1
                txtNum.Enabled = True
            Case 2
                txtNum.Enabled = False
                txtNum.Text = 1
                My.Settings.label_number = txtNum.Text
                My.Settings.Save()
            Case 3
                txtNum.Enabled = False
                txtNum.Text = 1
                My.Settings.label_number = txtNum.Text
                My.Settings.Save()
        End Select
    End Sub
End Class