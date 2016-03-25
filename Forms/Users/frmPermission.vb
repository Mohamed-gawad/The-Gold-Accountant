Public Class frmPermission
    Dim Myconn As New Connection
    Dim fin As Boolean
    Private Sub New_record()
        Myconn.ClearAllControls(GroupBox1, True)
    End Sub
    Private Sub Filldrg()
        If Not fin Then Return
        drg.Rows.Clear()
        Myconn.ExecQuery("SELECT  P.Employee_ID,P.U_Full,P.U_new,P.U_add,P.U_delete,P.U_Back,P.U_Updat,P.U_Search,P.U_Print,M.Menu_Text,S.Sub_menu_text,U.User_Password,U.Employee_Name,P.ID
                            FROM (( Users_Permission P LEFT JOIN Main_menu M  ON P.Main_menu_ID = M.Main_menu_ID) 
                            left join Sub_menu S on P.Sub_menu_ID =S.Sub_menu_ID)
                            left join Users_ID U on P.Employee_ID = U.Employee_ID
                             where p.Employee_ID = " & CInt(CboEmployee.SelectedValue) & " order by P.ID")

        If Not String.IsNullOrEmpty(Myconn.Exception) Then MsgBox(Myconn.Exception) : Exit Sub

        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("Employee_Name")
            drg.Rows(i).Cells(2).Value = r("Menu_Text")
            drg.Rows(i).Cells(3).Value = r("Sub_menu_text")
            drg.Rows(i).Cells(4).Value = r("U_Full")
            drg.Rows(i).Cells(5).Value = r("U_new")
            drg.Rows(i).Cells(6).Value = r("U_add")
            drg.Rows(i).Cells(7).Value = r("U_Updat")
            drg.Rows(i).Cells(8).Value = r("U_delete")
            drg.Rows(i).Cells(9).Value = r("U_Back")
            drg.Rows(i).Cells(10).Value = r("U_Search")
            drg.Rows(i).Cells(11).Value = r("U_Print")
            drg.Rows(i).Cells(12).Value = r("ID")
        Next
        Myconn.DataGridview_MoveLast(drg, 2)
    End Sub
    Private Sub Binding()
        Try
            Myconn.ExecQuery("SELECT * from Users_Permission where ID = " & CInt(drg.CurrentRow.Cells(12).Value) & " order by ID")
            Dim r As DataRow = Myconn.dt.Rows(0)
            CboEmployee.SelectedValue = r("Employee_ID")
            cbo_Main_menu.SelectedValue = r("Main_menu_ID")
            cbo_Sub_menu.SelectedValue = r("Sub_menu_ID")
            Check_Add.Checked = r("U_add")
            Check_Back.Checked = r("U_back")
            Check_Delet.Checked = r("U_delete")
            Check_Full.Checked = r("U_Full")
            Check_New.Checked = r("U_new")
            Check_Print.Checked = r("U_Print")
            Check_Search.Checked = r("U_Search")
            Check_Update.Checked = r("U_Updat")
        Catch ex As Exception
            Return
        End Try
    End Sub
    Private Sub Save_recod()
        With Myconn
            .Parames.Clear()
            .Addparam("@Employee_ID", CboEmployee.SelectedValue) ' المستخدم
            .Addparam("@U_Full", Check_Full.Checked) ' تحكم كامل
            .Addparam("@U_new", Check_New.Checked) ' جديد
            .Addparam("@U_add", Check_Add.Checked) ' حفظ
            .Addparam("@U_delete", Check_Delet.Checked) ' حذف
            .Addparam("@U_Back", Check_Back.Checked) ' مرتجع
            .Addparam("@U_Updat", Check_Update.Checked) ' تعديل
            .Addparam("@U_Search", Check_Search.Checked) ' بحث
            .Addparam("@U_Print", Check_Print.Checked) ' طباعة
            .Addparam("@Main_menu_ID", cbo_Main_menu.SelectedValue) ' قائمة رئيسية
            .Addparam("@Sub_menu_ID", If(cbo_Sub_menu.SelectedIndex = -1, DBNull.Value, cbo_Sub_menu.SelectedValue)) ' قائمة فرعية
            .ExecQuery("insert into  [Users_Permission] (Employee_ID,U_Full,U_new,U_add,U_delete,U_Back,U_Updat,U_Search,U_Print,Main_menu_ID,Sub_menu_ID) values(@Employee_ID,@U_Full,@U_new,@U_add,@U_delete,@U_Back,@U_Updat,@U_Search,@U_Print,@Main_menu_ID,@Sub_menu_ID)")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
        MessageBox.Show("تمت عملية الحفظ بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub
    Private Sub Update_record()
        With Myconn
            .Parames.Clear()
            .Addparam("@Employee_ID", CboEmployee.SelectedValue) ' المستخدم
            .Addparam("@U_Full", Check_Full.Checked) ' تحكم كامل
            .Addparam("@U_new", Check_New.Checked) ' جديد
            .Addparam("@U_add", Check_Add.Checked) ' حفظ
            .Addparam("@U_delete", Check_Delet.Checked) ' حذف
            .Addparam("@U_Back", Check_Back.Checked) ' مرتجع
            .Addparam("@U_Updat", Check_Update.Checked) ' تعديل
            .Addparam("@U_Search", Check_Search.Checked) ' بحث
            .Addparam("@U_Print", Check_Print.Checked) ' طباعة
            .Addparam("@Main_menu_ID", cbo_Main_menu.SelectedValue) ' قائمة رئيسية
            .Addparam("@Sub_menu_ID", If(cbo_Sub_menu.SelectedIndex = -1, DBNull.Value, cbo_Sub_menu.SelectedValue)) ' قائمة فرعية
            .Addparam("@ID", drg.CurrentRow.Cells(12).Value)
            .ExecQuery("Update [Users_Permission] set  Employee_ID=@Employee_ID,U_Full=@U_Full,U_new=@U_new,U_add=@U_add,U_delete=@U_delete,U_Back=@U_Back,U_Updat=@U_Updat,U_Search=@U_Search,U_Print=@U_Print,Main_menu_ID=@Main_menu_ID,Sub_menu_ID=@Sub_menu_ID where ID=@ID")
            If Myconn.NoErrors(True) = False Then Exit Sub
        End With
        MessageBox.Show("تمت عملية التعديل بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

    End Sub

    Private Sub frmPermission_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label5.Left = 0
            Label5.Width = Me.Width

            If F <> 1 Then
                Myconn.ExecQuery("Select * from Users_Permission where Employee_ID =" & CInt(My.Settings.user_ID) & " and Sub_menu_ID = " & Per & "")
                If Myconn.dt.Rows.Count = 0 Then GoTo A
                Dim r As DataRow = Myconn.dt.Rows(0)
                If r("U_full").ToString = False Then
                    btnSave.Enabled = r("U_add").ToString '
                    btnUpdat.Enabled = r("U_updat").ToString '
                    btnNew.Enabled = r("U_new").ToString '
                    btnDel.Enabled = r("U_delete").ToString '
                    btnPrint.Enabled = r("U_print").ToString '
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
A:
        Myconn.Fillcombo("Select * from Employees order by Employee_Name", "Employees", "Employee_ID", "Employee_Name", Me, CboEmployee)
        fin = False
        Myconn.Fillcombo("Select * from Main_menu", "Main_menu", "Main_menu_ID", "Menu_Text", Me, cbo_Main_menu)
        fin = True

    End Sub
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        New_record()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is ComboBox Then
                If txt.Text = "" And txt.Name <> "cbo_Sub_menu" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            End If
        Next
        Myconn.ExecQuery("select * from Users_Permission where Employee_ID =" & CInt(CboEmployee.SelectedValue) & " and Main_menu_ID =" & CInt(cbo_Main_menu.SelectedValue) & " and Sub_menu_ID =" & CInt(cbo_Sub_menu.SelectedValue))
        If Myconn.Recodcount > 0 Then MsgBox(" توجد صلاحيات لهذه القائمة يمكنك التعديل عليها", MsgBoxStyle.MsgBoxRtlReading & MsgBoxStyle.Critical, "رسالة") : Return
        Save_recod()
        Filldrg()

    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If MsgBox("هل أنت متأكد من عملية الحذف ؟", MsgBoxStyle.YesNo, "تأكيد الحذف") = MsgBoxResult.No Then Exit Sub
        With Myconn
            .Addparam("@ID", drg.CurrentRow.Cells(12).Value)
            .ExecQuery("delete from [Users_Permission] where ID = @ID ")
        End With
        If Myconn.NoErrors(True) = False Then Exit Sub
        Filldrg()

    End Sub

    Private Sub btnUpdat_Click(sender As Object, e As EventArgs) Handles btnUpdat.Click
        For Each txt As Control In GroupBox1.Controls
            If TypeOf txt Is ComboBox Then
                If txt.Text = "" Then
                    ErrorProvider1.SetError(txt, "أكمل البيانات")
                    MessageBox.Show("من فضلك أكمل البيانات ", "رسالة تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
                    Return
                End If
            End If
        Next

        Update_record()
        Filldrg()

    End Sub

    Private Sub CboEmployee_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboEmployee.SelectedIndexChanged
        Filldrg()

    End Sub

    Private Sub drg_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles drg.CellClick
        Binding()

    End Sub

    Private Sub cbo_Main_menu_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbo_Main_menu.SelectedValueChanged
        If Not fin Then Return
        cbo_Sub_menu.DataSource = Nothing
        Myconn.Fillcombo("Select * from Sub_menu where Main_menu_ID =" & CInt(cbo_Main_menu.SelectedValue), "Sub_menu", "Sub_menu_ID", "Sub_menu_text", Me, cbo_Sub_menu)
    End Sub

    Private Sub Check_Full_CheckedChanged(sender As Object, e As EventArgs) Handles Check_Full.CheckedChanged
        If Check_Full.Checked = False Then
            For Each txt As Control In GroupBox1.Controls
                If TypeOf txt Is CheckBox Then
                    If txt.Text <> "تحكم كامل" Then
                        txt.Enabled = True
                    End If
                End If
            Next
        Else
            For Each txt As Control In GroupBox1.Controls
                If TypeOf txt Is CheckBox Then
                    If txt.Text <> "تحكم كامل" Then
                        txt.Enabled = False
                    End If
                End If
            Next
        End If

    End Sub
End Class