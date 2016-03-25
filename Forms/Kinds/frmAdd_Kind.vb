Public Class frmAdd_Kind
    Dim Myconn As New Connection
    Dim fin As String

    Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)
        Myconn.Autonumber("items_Cod", "Items", txtID, Me)
        txtParcode.Text = Format(CInt(txtID.Text), "000000")
    End Sub
    Sub Save_record()
        With Myconn
            .Parames.Clear()
            .Addparam("@items_Cod", txtID.Text)
            .Addparam("@Items_Name", txtName.Text)
            .Addparam("@Group_cod", cbo_Group.SelectedValue)
            .Addparam("@Parcode", txtParcode.Text.Trim)
            .Addparam("@Customer_Price", DBNull.Value)
            .Addparam("@Total_Price", DBNull.Value)
            .Addparam("@cost_Price", DBNull.Value)
            .Addparam("@Supplier_ID", cbo_Supplier.SelectedValue)
            .ExecQuery("insert into  Items (items_Cod,Items_Name,Group_cod,Parcode,Customer_Price,Total_Price,cost_Price,Supplier_ID) 
                                     values(@items_Cod,@Items_Name,@Group_cod,@Parcode,@Customer_Price,@Total_Price,@cost_Price,@Supplier_ID)")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
    End Sub

    Private Sub Check_Parcode()
        Try
            Myconn.ExecQuery("Select * from Items where Parcode ='" & txtParcode.Text & "'")
            If Myconn.Recodcount = 0 Then Return
            Dim r As DataRow = Myconn.dt.Rows(0)
            If Myconn.Recodcount > 0 Then
                MsgBox("  الباركود " & r("Parcode").ToString & " مسجل باسم الصنف" & vbNewLine & r("Items_name").ToString, MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة")
                txtParcode.Text = ""
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical & MsgBoxStyle.MsgBoxRtlReading, "رسالة")
        End Try
    End Sub
    Private Sub frmAdd_Kind_Load(sender As Object, e As EventArgs) Handles Me.Load
        Label5.Left = 0
        Label5.Width = Me.Width
        fin = False
        Myconn.Fillcombo("select * from [group] order by [group_Name]", "[group]", "Group_cod", "group_Name", Me, cbo_Group)
        Myconn.Fillcombo("select * from Supplier order by Supplier_Name", "Supplier", "Supplier_ID", "Supplier_Name", Me, cbo_Supplier)
        fin = True
        New_record()
        Myconn.Autocomplete("items", "items_name", txtName)

        '-------------------------------------------------------------------------------------------------- النسخة التجريبية
        'Myconn.ExecQuery("select * from Items")
        'If Myconn.Recodcount > 20 Then
        '    MsgBox("هذه النسخة تجريبية")
        '    btnSave.Enabled = False
        '    btnNew.Enabled = False
        '    Return
        'End If
    End Sub
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        New_record()
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Check_Parcode()
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is TextBox Then
                If txt.Text = "" Then
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
        New_record()
        Myconn.Autocomplete("items", "items_name", txtName)
    End Sub
    Private Sub txtParcode_Leave(sender As Object, e As EventArgs) Handles txtParcode.Leave
        Check_Parcode()
    End Sub
    Private Sub langAR_Enter(sender As Object, e As EventArgs) Handles txtName.Enter, cbo_Group.Enter, cbo_Supplier.Enter
        Myconn.langAR()
    End Sub
    Private Sub NumberOnly_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtParcode.KeyPress
        Myconn.NumberOnly(txtParcode, e)
    End Sub
    Private Sub Error_Clear(sender As Object, e As EventArgs) Handles txtName.TextChanged
        ErrorProvider1.Clear()
    End Sub

    Private Sub txtParcode_TextChanged(sender As Object, e As EventArgs) Handles txtParcode.TextChanged
        ErrorProvider1.Clear()
        LBarcode.Text = "*" & txtParcode.Text.Trim & "*"
    End Sub

    Private Sub btnAdd_group_Click(sender As Object, e As EventArgs) Handles btnAdd_group.Click
        Try
            Dim frm As New frmAdd_Group
            frm.MdiParent = FrmMain
            frm.Show()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub frmAdd_Kind_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        cbo_Group.DataSource = Nothing
        cbo_Supplier.DataSource = Nothing
        Myconn.Fillcombo("select * from [group] order by [group_Name]", "[group]", "Group_cod", "group_Name", Me, cbo_Group)
        Myconn.Fillcombo("select * from Supplier order by Supplier_Name", "Supplier", "Supplier_ID", "Supplier_Name", Me, cbo_Supplier)

    End Sub

    Private Sub btn_Supplier_Click(sender As Object, e As EventArgs) Handles btn_Supplier.Click
        Try
            Dim frm As New frmAdd_Supplier
            frm.MdiParent = FrmMain
            frm.Show()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txtParcode_KeyUp(sender As Object, e As KeyEventArgs) Handles txtParcode.KeyUp
        If e.KeyCode = Keys.Enter Then
            If btnSave.Enabled = False Then MsgBox("هذه النسخة تجريبية", MsgBoxStyle.Critical, "رسالة") : Return

            btnSave_Click(Nothing, Nothing)
            txtID.Focus()

        End If
    End Sub

#Region "Moving"
    Private Sub txtName_KeyUp(sender As Object, e As KeyEventArgs) Handles txtName.KeyUp
        If e.KeyCode = Keys.Enter = True Then
            cbo_Group.Focus()
        End If
    End Sub
    Private Sub cbo_Group_KeyUp(sender As Object, e As KeyEventArgs) Handles cbo_Group.KeyUp
        If e.KeyCode = Keys.Enter = True Then
            cbo_Supplier.Focus()
        End If
    End Sub
    Private Sub cbo_Supplier_KeyUp(sender As Object, e As KeyEventArgs) Handles cbo_Supplier.KeyUp
        If e.KeyCode = Keys.Enter = True Then
            txtParcode.Focus()
        End If
    End Sub

    Private Sub txtID_KeyUp(sender As Object, e As KeyEventArgs) Handles txtID.KeyUp
        If e.KeyCode = Keys.Enter = True Then
            txtName.Focus()
        End If
    End Sub

#End Region


End Class