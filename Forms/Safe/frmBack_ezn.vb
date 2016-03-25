Imports System.Globalization
Public Class frmBack_ezn
    Dim fin As Boolean
    Dim Myconn As New Connection
    Dim St1, St2 As String
    Dim x, y As Integer
    Dim S, W As Double



    Private Sub Fillgrd()
        drg.Rows.Clear()

        Select Case cbo_ezn.SelectedIndex
            Case 0 ' إذن دفع
                Myconn.ExecQuery("Select Y.perm_ID,Y.per_ID,Y.per_date,Y.Amount,Y.Note_per,Y.Status,S.Employee_Name,F.perm_name,(e.Supplier_Name) as Customer_Name,(t.pay_Item_name) as Recive_Item_name
                            From (((Safe_payment_per Y Left join Users_ID S on Y.users_ID = S.Employee_ID )
                            Left join Safe_Per F on Y.perm_ID = F.perm_ID)
                            Left join Supplier e on Y.Supplier_ID = e.Supplier_ID)
                            Left join Pay_Items t on Y.pay_Item_ID = t.pay_Item_ID " & St2 & " order by Y.per_date")

            Case 1 ' إذن استلام
                Myconn.ExecQuery("Select R.perm_ID, R.per_ID,R.per_date,R.Amount,R.Note_per,R.Status ,U.Employee_Name,P.perm_name,C.Customer_Name,I.Recive_Item_name
                            From (((Safe_recive_per R Left join Users_ID U on R.users_ID = U.Employee_ID )
                            Left join Safe_Per P on R.perm_ID = P.perm_ID)
                            Left join Customers C on R.Customer_ID = C.customer_ID)
                            Left join Recive_Items I on R.Recive_Item_ID = I.Recive_Item_ID " & St1 & " order by R.per_date")
            Case 2 '  جميع الأذونات
                Myconn.ExecQuery("Select R.perm_ID, R.per_ID, R.per_date, R.Amount, R.Note_per, R.Status, U.Employee_Name, P.perm_name, C.Customer_Name, I.Recive_Item_name
                            From (((Safe_recive_per R Left join Users_ID U on R.users_ID = U.Employee_ID )
                            Left join Safe_Per P on R.perm_ID = P.perm_ID)
                            Left join Customers C on R.Customer_ID = C.customer_ID)
                            Left join Recive_Items I on R.Recive_Item_ID = I.Recive_Item_ID " & St1 & "
                            UNION ALL
                            Select Y.perm_ID,Y.per_ID,Y.per_date,Y.Amount,Y.Note_per,Y.Status,S.Employee_Name,F.perm_name,e.Supplier_Name,t.pay_Item_name
                            From (((Safe_payment_per Y Left join Users_ID S on Y.users_ID = S.Employee_ID )
                            Left join Safe_Per F on Y.perm_ID = F.perm_ID)
                            Left join Supplier e on Y.Supplier_ID = e.Supplier_ID)
                            Left join Pay_Items t on Y.pay_Item_ID = t.pay_Item_ID " & St2 & " order by R.per_date")
        End Select
        W = 0
        S = 0
        For i As Integer = 0 To Myconn.dt.Rows.Count - 1
            Dim r As DataRow = Myconn.dt.Rows(i)
            drg.Rows.Add()
            drg.Rows(i).Cells(0).Value = i + 1
            drg.Rows(i).Cells(1).Value = r("perm_name")
            drg.Rows(i).Cells(2).Value = r("Recive_Item_name")
            drg.Rows(i).Cells(3).Value = Format(CDate(r("per_date")), "yyyy/MM/dd")
            drg.Rows(i).Cells(4).Value = If(IsDBNull(r("Customer_Name")), r("Note_per"), r("Customer_Name"))
            drg.Rows(i).Cells(5).Value = r("amount")
            drg.Rows(i).Cells(6).Value = clsNumber.nTOword(r("amount"))
            drg.Rows(i).Cells(7).Value = r("Employee_Name")
            drg.Rows(i).Cells(8).Value = r("Note_per")
            drg.Rows(i).Cells(9).Value = r("Status")
            If r("perm_ID") = 1 Then
                drg.Rows(i).DefaultCellStyle.BackColor = Color.LemonChiffon
                W += r("amount")
            Else
                drg.Rows(i).DefaultCellStyle.BackColor = Color.Pink
                S += r("amount")
            End If
        Next
        Label4.Text = W
        Label6.Text = clsNumber.nTOword(Label4.Text)
        Label6.Left = Label4.Left - (Label6.Width + 20)
        Label7.Text = S
        Label8.Text = clsNumber.nTOword(Label7.Text)
        Label8.Left = Label7.Left - (Label8.Width + 20)

    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click

        'MsgBox(Format(CDate(txt1.Text), "yyyy/MM/dd"))
    End Sub

    Private Sub frmBack_ezn_Load(sender As Object, e As EventArgs) Handles Me.Load
        Label5.Left = 0
        Label5.Width = Me.Width

    End Sub
    Private Sub cbo_ezn_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_ezn.SelectedIndexChanged
        Select Case cbo_ezn.SelectedIndex
            Case 0 ' اذن دفع
                St2 = " where Y.perm_ID =2 and Y.Status = False"
            Case 1 ' اذن استلام
                St1 = " where R.perm_ID =1 and R.Status = false"

            Case 2 ' كل الاذونات
                St1 = " where  R.Status = false"
                St2 = " where  Y.Status = False"
        End Select
        Fillgrd()
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try


            If txt1.Text = "" Or txt2.Text = "" Then
                cbo_ezn_SelectedIndexChanged(Nothing, Nothing)
                Return
            End If
            Select Case cbo_ezn.SelectedIndex
                Case 0 ' إذن دفع
                    St2 = " where  Y.Status = False and Y.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"

                Case 1 ' إذن استلام
                    St1 = " where R.Status = false and R.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"

                Case 2 '  جميع الأذونات
                    St1 = " where R.Status = false and R.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"
                    St2 = " where Y.Status = False and Y.per_date between #" & Format(CDate(txt1.Text), "yyyy/MM/dd").ToString & "# And #" & Format(CDate(txt2.Text), "yyyy/MM/dd").ToString & "#"
            End Select
            Fillgrd()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class