Public Class frmBack_Purchases
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim st As String
    Private Sub Fillgrd()
        drg.Rows.Clear()
        Select Case cboSearch.SelectedIndex
            Case 0 '  مرتجعات صنف
                st = " and Purchases.items_Cod = " & CInt(cboSupplier.ComboBox.SelectedValue)
            Case 1 '   صنف خلال فترة
                st = " and Purchases.items_Cod = " & CInt(cboSupplier.ComboBox.SelectedValue) & " and Purchases.Pur_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
            Case 2 '  مرتجعات خلال فترة
                st = " and Purchases.Pur_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
            Case 3 '   مرتجعات مورد
                st = " and Purchases.Supplier_ID = " & CInt(cboSupplier.ComboBox.SelectedValue)
            Case 4 '   مرتجعات مورد خلال فترة
                st = " and Purchases.Supplier_ID = " & CInt(cboSupplier.ComboBox.SelectedValue) & " and Purchases.Pur_Date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd") & "# and #" & Format(CDate(txt2.Text), "yyyy/MM/dd") & "#"
            Case 5 '  كل المرتجعات
                st = Nothing
        End Select
        Myconn.ExecQuery("SELECT Purchases.Status AS Expr1, Purchases.ID, Purchases.Pur_Date, Purchases.Pur_Time, Purchases.Pur_Bill_num, Purchases.Supplier_ID, Purchases.items_Cod, Purchases.Customer_Price, Purchases.Reduce, Purchases.Pur_Price, Purchases.Iteme_Number, Purchases.Total_Price, Purchases.Stock_ID, Purchases.Status, Items.Items_Name,Items.Parcode, Employees.Employee_Name, Supplier.Supplier_Name
                                    FROM Supplier RIGHT JOIN ((Employees RIGHT JOIN Users_ID ON Employees.Employee_ID = Users_ID.Employee_ID) 
                                    RIGHT JOIN (Purchases LEFT JOIN Items ON Purchases.items_Cod = Items.items_Cod) ON Users_ID.Employee_ID = Purchases.Employee_ID) ON Supplier.Supplier_ID = Purchases.Supplier_ID
                                    WHERE (((Purchases.Status)=False))" & st & " order by Purchases.ID")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub
        Dim V1 As Double = 0
        Dim V2 As Double = 0
        Dim B As Double = 0
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1 ' المسلسل
            drg.Rows(i).Cells(1).Value = Format(CDate(r("Pur_Date")), "yyyy/MM/dd") ' التاريخ
            drg.Rows(i).Cells(2).Value = r("Pur_Bill_num") ' رقم الفاتورة
            drg.Rows(i).Cells(3).Value = r("Supplier_Name") ' المورد
            drg.Rows(i).Cells(4).Value = r("Items_Name") ' الصنف
            drg.Rows(i).Cells(5).Value = r("Parcode") ' باركود الصنف
            drg.Rows(i).Cells(6).Value = r("Customer_Price") ' سعر المصنع
            drg.Rows(i).Cells(7).Value = r("Reduce") & " % " ' الخصم
            drg.Rows(i).Cells(8).Value = r("Pur_Price") ' التكلفة
            drg.Rows(i).Cells(9).Value = r("Iteme_Number") ' العدد
            drg.Rows(i).Cells(10).Value = r("Customer_Price") * r("Iteme_Number") ' اجمالي المصنع
            drg.Rows(i).Cells(11).Value = r("Total_Price") ' صافي الاجمالي
            drg.Rows(i).Cells(12).Value = r("Employee_Name") ' المستخدم

            V1 += CDec(drg.Rows(i).Cells(11).Value) ' صافي التكلفة
            V2 += CDec(drg.Rows(i).Cells(10).Value) ' اجمالي المصنع
        Next

        Myconn.DataGridview_MoveLast(drg, 2)
        Label19.Text = V2 ' اجمالي المصنع
        Label18.Text = V1 ' اجمالي التكلفة
        Label20.Text = Math.Round(((Val(Val(Label19.Text) - Val(Label18.Text)) / Val(Label19.Text)) * 100), 2) & " % " ' نسبة الخصم
    End Sub

    Private Sub frmBack_Purchases_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width

            If  F <> 1Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then MsgBox("قم باضافة المستخدمين واضافة صلاحيات للتعامل مع هذه النافذة", MsgBoxStyle.Critical, "رسالة تنبيه") : Exit Sub
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnSearch.Enabled = r("U_search").ToString
                    btnPrint.Enabled = r("U_print").ToString
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        fillgrd()
    End Sub

    Private Sub cboSearch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSearch.SelectedIndexChanged
        Select Case cboSearch.SelectedIndex
            Case 0 '  مرتجعات صنف
                L1.Visible = False
                L2.Visible = False
                txt1.Visible = False
                txt2.Visible = False
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("select Items_Name,items_Cod from [items] order by [Items_Name]", "[items]", "items_Cod", "Items_Name", Me, cboSupplier.ComboBox)
            Case 1 '   صنف خلال فترة
                L1.Visible = True
                L2.Visible = True
                txt1.Visible = True
                txt2.Visible = True
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("select Items_Name,items_Cod from [items] order by [Items_Name]", "[items]", "items_Cod", "Items_Name", Me, cboSupplier.ComboBox)

            Case 2 '  مرتجعات خلال فترة
                L1.Visible = True
                L2.Visible = True
                txt1.Visible = True
                txt2.Visible = True
                cboSupplier.Visible = False

            Case 3 '   مرتجعات مورد
                L1.Visible = False
                L2.Visible = False
                txt1.Visible = False
                txt2.Visible = False
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("select * from Supplier order by Supplier_Name", "Supplier", "Supplier_ID", "Supplier_Name", Me, cboSupplier.ComboBox)
            Case 4 '   مرتجعات مورد خلال فترة
                L1.Visible = True
                L2.Visible = True
                txt1.Visible = True
                txt2.Visible = True
                cboSupplier.Visible = True
                cboSupplier.ComboBox.DataSource = Nothing
                Myconn.Fillcombo("select * from Supplier order by Supplier_Name", "Supplier", "Supplier_ID", "Supplier_Name", Me, cboSupplier.ComboBox)

            Case 5 '  كل المرتجعات
                L1.Visible = False
                L2.Visible = False
                txt1.Visible = False
                txt2.Visible = False
                cboSupplier.Visible = False
        End Select
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Print_bills()
    End Sub
    Private Sub Print_bills()
        Dim table As New DataTable
        For i As Integer = 1 To 11
            Dim x As String
            x = Format(i, "00")
            table.Columns.Add(x)
        Next

        For Each dr As DataGridViewRow In drg.Rows
            table.Rows.Add()
            table.Rows(table.Rows.Count - 1)(0) = table.Rows.Count
            table.Rows(table.Rows.Count - 1)(1) = dr.Cells(2).Value ' رقم الفاتورة
            table.Rows(table.Rows.Count - 1)(2) = dr.Cells(1).Value ' التاريخ
            table.Rows(table.Rows.Count - 1)(3) = dr.Cells(3).Value ' المورد
            table.Rows(table.Rows.Count - 1)(4) = dr.Cells(4).Value ' الصنف
            table.Rows(table.Rows.Count - 1)(5) = dr.Cells(5).Value ' الباركود
            table.Rows(table.Rows.Count - 1)(6) = dr.Cells(6).Value ' سعر المصنع 
            table.Rows(table.Rows.Count - 1)(7) = dr.Cells(9).Value ' العدد
            table.Rows(table.Rows.Count - 1)(8) = dr.Cells(10).Value ' الاجمالي
            table.Rows(table.Rows.Count - 1)(9) = dr.Cells(7).Value ' الخصم
            table.Rows(table.Rows.Count - 1)(10) = dr.Cells(11).Value ' سعر التكلفة بعد الخصم
        Next
        Dim rpt As New rpt_Pur_kind_Back
        rpt.SetDataSource(table)
        rpt.SetParameterValue("Co", My.Settings.Co_name)
        rpt.SetParameterValue("Address", "العنوان : " & My.Settings.Co_address & " تليفون : " & My.Settings.Co_tel)
        'rpt.SetParameterValue("Bill_num", Format(CDate(txt1.Text), "yyyy/MM/dd"))
        rpt.SetParameterValue("Price", Label18.Text)
        rpt.SetParameterValue("reduce", Math.Round((Label18.Text - Label19.Text), 2))
        rpt.SetParameterValue("Nisba", Label20.Text)
        rpt.SetParameterValue("Total", Label19.Text)


        If My.Settings.Print = True Then
            frmReportViewer.CrystalReportViewer1.ReportSource = rpt
            frmReportViewer.Show()
        Else
            rpt.PrintOptions.PrinterName = My.Settings.Printer_report
            rpt.PrintToPrinter(1, False, 0, 0)
        End If

    End Sub

End Class