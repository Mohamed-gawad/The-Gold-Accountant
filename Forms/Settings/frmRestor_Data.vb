Imports System.IO
Public Class frmRestor_Data
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Using openDialog As New OpenFileDialog()
                openDialog.CheckFileExists = True
                openDialog.CheckPathExists = True
                openDialog.Filter = "Microsoft Access Database (*.accdb)|*.accdb"
                openDialog.RestoreDirectory = True
                If openDialog.ShowDialog() = DialogResult.OK Then
                    If File.Exists(openDialog.FileName) Then
                        File.Delete(Application.StartupPath & "\Data.accdb")
                        File.Copy(openDialog.FileName, Application.StartupPath & " \Data.accdb")
                    End If
                End If
            End Using
            MessageBox.Show("تم استعادة قاعدة البيانات بنجاح", "رسالة", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
        Catch ex As Exception
            MsgBox("يجب إغلاق جميع نوافذ البرنامج أثناء عملية الإصلاح ", MsgBoxStyle.Critical, "تنبيه")
        End Try
    End Sub
End Class