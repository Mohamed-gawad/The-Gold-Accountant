Imports System.Drawing.Printing
Public Class frmSetting
    Dim myconn As New Connection
    Private Sub frmSetting_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Myconn.Fillcombo("select * from [Stocks] order by Stock_ID", "[Stocks]", "Stock_ID", "Stock_Name", Me, cbo_Stock)


            Dim pkInstalledPrinters As String

            ' Find all printers installed
            For Each pkInstalledPrinters In
                PrinterSettings.InstalledPrinters
                cboPrinters.Items.Add(pkInstalledPrinters)
                cboPrinter_report.Items.Add(pkInstalledPrinters)
            Next pkInstalledPrinters
            CheckBox1.Checked = My.Settings.Price_sales
            cboPrinter_report.Text = My.Settings.Printer_report
            cboPrinters.Text = My.Settings.Printer_Sales
            CheckBox2.Checked = My.Settings.Reduce
            TextBox1.Text = My.Settings.Reduce_amount
            CheckBox3.Checked = If(My.Settings.Sales_Case, False, True)
            cboPrice.SelectedIndex = My.Settings.Sales_Kind
            CheckBox4.Checked = If(My.Settings.S_Stock, False, True)
            cbo_Stock.SelectedValue = My.Settings.Stock_ID
            CheckBox5.Checked = My.Settings.Print
            Check_Factory.Checked = My.Settings.Factory_Price
            Check_Customer_Price.Checked = My.Settings.Customer_Price
            cbo_Sales_bill.SelectedIndex = My.Settings.Sales_Bill_Kind
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Settings.Printer_Sales = cboPrinters.Text
        My.Settings.Printer_report = cboPrinter_report.Text
        My.Settings.Price_sales = CheckBox1.Checked
        My.Settings.Reduce = CheckBox2.Checked
        My.Settings.Reduce_amount = If(TextBox1.Text = "", 0, TextBox1.Text)
        My.Settings.Sales_Case = If(CheckBox3.Checked = True, False, True)
        My.Settings.Sales_Kind = cboPrice.SelectedIndex
        My.Settings.S_Stock = If(CheckBox4.Checked = True, False, True)
        My.Settings.Stock_ID = cbo_Stock.SelectedValue
        My.Settings.Print = CheckBox5.Checked
        My.Settings.Factory_Price = Check_Factory.Checked
        My.Settings.Customer_Price = Check_Customer_Price.Checked
        My.Settings.Sales_Bill_Kind = cbo_Sales_bill.SelectedIndex
        My.Settings.Save()

        MessageBox.Show("تم حفظ الاعدادات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub
End Class