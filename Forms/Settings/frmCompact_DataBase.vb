Imports System.IO
Public Class frmCompact_DataBase
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim dbe As New Microsoft.Office.Interop.Access.Dao.DBEngine
            dbe.CompactDatabase(Application.StartupPath & "\Data.accdb", Application.StartupPath & "\Data_Compact.accdb", , , ";pwd=moh611974182008")
            File.Delete(Application.StartupPath & "\Data.accdb")
            Rename(Application.StartupPath & "\Data_Compact.accdb", Application.StartupPath & "\Data.accdb")
            MessageBox.Show("تم ضغط وإصلاح قاعدة البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)

        Catch ex As Exception
            MsgBox("يجب إغلاق جميع نوافذ البرنامج أثناء عملية الإصلاح ", MsgBoxStyle.Critical, "تنبيه")
        End Try

    End Sub
End Class