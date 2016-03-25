Imports System.Management
Public Class frmLogin
    Dim myconn As New Connection
    Dim fin As Boolean
    Dim A, P As String
    Public Function GetDriveSerialNumber() As String ' فحص رقم الهارد
        Dim HDD_Serial, HDD_ModelNumber As String
        Dim hdd As New ManagementObjectSearcher("select * from Win32_DiskDrive")

        For Each hd In hdd.Get
            HDD_Serial = hd("SerialNumber").ToString.Trim
            HDD_ModelNumber = hd("Model").ToString.Trim

        Next
        Return HDD_Serial & HDD_ModelNumber
    End Function
    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles Me.Load

        '-------------------------------------------------------------------------------------------------------------- تصريح استخدام البرنامج

        If GetDriveSerialNumber() <> Label3.Text Then
            MsgBox(" هذه النسخة غير مصرح باستخدامها قم بالاتصال بمصمم البرنامج  " & vbNewLine & "على رقم 01125139439 لشراء حق استخدام البرنامج                            ", MsgBoxStyle.MsgBoxRtlReading, "رسالة ")
            Close()

        End If

        '------------------------------------------------------------------------------------------------------------------------------------------------------

        myconn.ExecQuery("Select * from Users_ID order by Employee_Name")
        If myconn.dt.Rows.Count = 0 Then
            MsgBox("مرحبا بكم هذا أول استخدام للبرنامج قم بتسجيل مدير النظام أولا", MsgBoxStyle.MsgBoxRight, "رسالة  ")
            F = 0
            FrmMain.Show()
            Exit Sub
        Else
            myconn.ExecQuery("Select * from Users_Permission ")
            If myconn.dt.Rows.Count = 0 Then
                MsgBox("قم باضافة  صلاحيات للمستخدمين", MsgBoxStyle.Critical, "رسالة تنبيه")
                F = 0
                FrmMain.Show()
                Exit Sub
            End If

            fin = False
            myconn.Fillcombo("Select * from Users_ID order by Employee_Name", "Users_ID", "Employee_ID", "Employee_Name", Me, cboEmployees)
            fin = True
        End If
    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click
        Dim Main_menu As ToolStripMenuItem
        Dim Sub_menu As ToolStripMenuItem
        If cboEmployees.SelectedIndex = -1 Then MsgBox("من فضلك قم باختيار اسم المستخدم من القائمة", MsgBoxStyle.MsgBoxRight, "رسالة") : Return
        If txtPass.Text = "" Then
            MsgBox("أدخل كلمة المرور", MsgBoxStyle.MsgBoxRight, "رسالة")
            Return
        ElseIf txtPass.Text <> P Then
            MsgBox("كلمة المرور غير صحيحة ...", MsgBoxStyle.MsgBoxRight, "رسالة")
            txtPass.Text = ""
            Return
        End If

        If txtPass.Text = P Then
            If F <> 1 Then
                For Each Main_menu In FrmMain.MenuStrip.Items
                    Main_menu.Visible = False
                    For Each Sub_menu In Main_menu.DropDownItems
                        Sub_menu.Visible = False
                    Next
                Next

                For i As Integer = 0 To myconn.dt.Rows.Count - 1
                    Dim r As DataRow = myconn.dt.Rows(i)
                    For Each Main_menu In FrmMain.MenuStrip.Items
                        If Main_menu.Name = r("Main_menu_name").ToString Then
                            Main_menu.Visible = True
                            For Each Sub_menu In Main_menu.DropDownItems
                                If Sub_menu.Name = r("Sub_menu_name").ToString Then
                                    Sub_menu.Visible = True
                                End If
                            Next
                        End If
                    Next

                Next

            End If
        End If
        My.Settings.user_ID = cboEmployees.SelectedValue
        My.Settings.Save()
        FrmMain.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()

    End Sub
    Private Sub cboEmployees_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboEmployees.SelectedIndexChanged
        If Not fin Then Return
        myconn.ExecQuery("SELECT P.Employee_ID,P.U_Full,P.U_new,P.U_add,P.U_delete,P.U_Back,P.U_Updat,P.U_Search,P.U_Print,M.Main_menu_name,S.Sub_menu_name,U.User_Password,U.job_ID
                            FROM (( Users_Permission P LEFT JOIN Main_menu M  ON P.Main_menu_ID = M.Main_menu_ID) 
                            left join Sub_menu S on P.Sub_menu_ID =S.Sub_menu_ID)
                            left join Users_ID U on P.Employee_ID = U.Employee_ID
                            where P.Employee_ID = " & CInt(cboEmployees.SelectedValue) & "")

        If myconn.dt.Rows.Count = 0 Then
            MsgBox(" عفوا ليس لك صلاحيات للدخول", MsgBoxStyle.MsgBoxRight, "تحذير")
            Return
        Else
            Dim r As DataRow = myconn.dt.Rows(0)
            P = r("User_Password").ToString
            F = r("job_ID")
        End If
        'MsgBox(F)
    End Sub

    Private Sub txtPass_KeyUp(sender As Object, e As KeyEventArgs) Handles txtPass.KeyUp
        If e.KeyCode = Keys.Enter Then
            btn_OK_Click(Nothing, Nothing)
        End If

    End Sub
End Class