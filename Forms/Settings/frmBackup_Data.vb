Imports System.IO
Imports System.Globalization
Public Class frmBackup_Data

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim dbe As New Microsoft.Office.Interop.Access.Dao.DBEngine
            If (Not Directory.Exists(Application.StartupPath & "\Backup")) Then
                Directory.CreateDirectory(Application.StartupPath & "\Backup")
            End If
        dbe.CompactDatabase(Application.StartupPath & "\Data.accdb", Application.StartupPath & "\Backup\" & Date.Now.ToString("yyyy-MM-dd") & Space(1) & TimeOfDay.ToString("hh mm ss tt", CultureInfo.CreateSpecificCulture("ar-eg")) & ".accdb", , , ";pwd=moh611974182008")
        MessageBox.Show("تم أخذ نسخة إحتياطية من قاعدة البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

        Catch ex As Exception
        MsgBox("يجب إغلاق جميع نوافذ البرنامج أثناء عملية أخذ نسخة إحتياطية من قاعدة البيانات ", MsgBoxStyle.Critical, "تنبيه")
        End Try

    End Sub
End Class