Imports System.Data.OleDb
Public Class Connection
    Public conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\Data.accdb;Jet OLEDB:Database Password = moh611974182008;")
    Public ds As New DataSet
    Public da As New OleDbDataAdapter
    Public dv As New DataView
    Public cmd As New OleDbCommand
    Public dr As OleDbDataReader
    Public dt As New DataTable
    Public cur As CurrencyManager
    Public Parames As New List(Of OleDbParameter)
    Public Recodcount As Integer
    Public Exception As String
    Sub Filldataset(SQl As String, tableName As String, frm As Form)
        ds = New DataSet
        Try
            da = New OleDbDataAdapter(SQl, conn)
            da.Fill(ds, tableName)
            dv = New DataView(ds.Tables(tableName))
            cur = CType(frm.BindingContext(dv), CurrencyManager)
        Catch ex As OleDbException
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub ExecQuery(SQL As String)
        Recodcount = 0
        Exception = ""
        Try
            If conn.State = ConnectionState.Open Then conn.Close()
            conn.Open()
            cmd = New OleDbCommand(SQL, conn)
            ' Load param to SQL command
            Parames.ForEach(Sub(x) cmd.Parameters.Add(x))
            ' Clear parmes List
            Parames.Clear()
            dt = New DataTable
            da = New OleDbDataAdapter(cmd)
            Recodcount = da.Fill(dt)
            conn.Close()
        Catch ex As Exception
            Exception = ex.Message
        End Try
        If conn.State = ConnectionState.Open Then conn.Close()
    End Sub
    Public Sub Addparam(Name As String, value As Object)
        Dim Newparam As New OleDbParameter(Name, value)
        Parames.Add(Newparam)
    End Sub
    Sub Fillcombo(SQL As String, tableName As String, ValueMember As String, DisplayMember As String, frm As Form, cbo As ComboBox)
        ExecQuery(SQL)
        cbo.Items.Clear()

        With cbo
            .DataSource = dt
            .ValueMember = ValueMember
            .DisplayMember = DisplayMember
            .SelectedValue = 0
        End With
    End Sub
    Sub TextBindingdata(frm As Form, grb As GroupBox, Fields() As String, txt() As TextBox)
        For i As Integer = 0 To Fields.Count - 1
            txt(i).DataBindings.Clear()
            txt(i).DataBindings.Add("text", dt, Fields(i))
        Next
    End Sub
    Sub Autonumber(col As String, table As String, txt As TextBox, frm As Form)
        Dim SQL As String
        SQL = "select Top 1 " & col & " from " & table & " order by " & col & " desc"
        ExecQuery(SQL)

        If dt.Rows.Count = 0 Then
            txt.Text = "1"
        Else
            Dim r As DataRow = dt.Rows(0)
            txt.Text = (r(col) + 1).ToString
        End If
    End Sub
    Sub langAR()
        Dim lan As New System.Globalization.CultureInfo("ar-eg")
        InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(lan)
    End Sub
    Sub langEN()
        Dim lan As New System.Globalization.CultureInfo("en-us")
        InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(lan)

    End Sub
    Sub NumberOnly(x As TextBox, e As KeyPressEventArgs)
        If Not Double.TryParse((x.Text & e.KeyChar), Nothing) Then
            e.Handled = True
        End If
        If Char.IsControl(e.KeyChar) Then
            e.Handled = False
        End If
        If e.KeyChar = "." Then
            e.Handled = False
        End If
    End Sub
    Sub Arabiconly(e As KeyPressEventArgs)

        Select Case e.KeyChar
            Case "ء" To "ي", ControlChars.Back, Chr(Keys.Space)
                e.Handled = False
            Case Else
                e.Handled = True
                MsgBox("الكتابة باللغة العربية")
        End Select
        langAR()
    End Sub
    Sub EnglishOnly(e As KeyPressEventArgs)
        Select Case e.KeyChar
            Case "A" To "z", ControlChars.Back, Chr(Keys.Space), "@", "."
                e.Handled = False
            Case "0" To "9"""
                e.Handled = False
            Case Else

                e.Handled = True
                MsgBox("الكتابة باللغة الإنجليزية")
        End Select
        langEN()
    End Sub
    Sub Autocomplete(table As String, col As String, txt As TextBox)
        Dim SQL As String
        SQL = "select  " & col & " from " & table & ""
        With conn
            If .State = ConnectionState.Closed Then
                .Open()
            End If
        End With
        cmd = New OleDbCommand(SQL, conn)
        dr = cmd.ExecuteReader()
        Dim Autocomp As New AutoCompleteStringCollection()
        While dr.Read()
            Autocomp.Add(dr(col))
        End While
        dr.Close()
        conn.Close()
        txt.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        txt.AutoCompleteSource = AutoCompleteSource.CustomSource
        txt.AutoCompleteCustomSource = Autocomp
    End Sub
    Sub comboBinding(dm As String, cbo As ComboBox)
        cbo.DataBindings.Clear()
        cbo.DataBindings.Add("SelectedValue", dv, dm)
    End Sub
    Sub Sum_drg(drg As DataGridView, y As Integer, B1 As Label, B2 As Label)
        Dim B As Decimal
        For i As Integer = 0 To drg.RowCount - 1
            B += CDec(drg.Rows(i).Cells(y).Value)
        Next
        B1.Text = B
        B2.Text = "(  " & clsNumber.nTOword(B1.Text) & "  )"
        B2.Left = (B1.Left - B2.Width) - 20
    End Sub
    Sub Sum_drg2(drg As DataGridView, y As Integer, B1 As Label)
        Dim B As Decimal
        For i As Integer = 0 To drg.RowCount - 1
            B += CDec(drg.Rows(i).Cells(y).Value)
        Next
        B1.Text = B

    End Sub
    Sub DataGridview_MoveLast(drg As DataGridView, x As Integer)
        If drg.Rows.Count = 0 Then Return
        drg.Rows(drg.Rows.Count - 1).Cells(x).Selected = True
        drg.CurrentCell = drg.SelectedCells(x)
    End Sub
    Sub ClearAllControls(container As GroupBox, Optional Recurse As Boolean = True)
        'Clear all of the controls within the container object
        'If "Recurse" is true, then also clear controls within any sub-containers
        Dim ctrl As Control
        For Each ctrl In container.Controls
            If (ctrl.GetType() Is GetType(TextBox)) Then
                Dim txt As TextBox = CType(ctrl, TextBox)
                txt.Text = ""
            End If
            If (ctrl.GetType() Is GetType(CheckBox)) Then
                Dim chkbx As CheckBox = CType(ctrl, CheckBox)
                chkbx.Checked = False
            End If
            If (ctrl.GetType() Is GetType(ComboBox)) Then
                Dim cbobx As ComboBox = CType(ctrl, ComboBox)
                cbobx.SelectedIndex = -1
            End If
            If (ctrl.GetType() Is GetType(DateTimePicker)) Then
                Dim dtp As DateTimePicker = CType(ctrl, DateTimePicker)
                dtp.Value = Now()
            End If

            If Recurse Then
                'If (ctrl.GetType() Is GetType(Panel)) Then
                '    Dim pnl As Panel = CType(ctrl, Panel)
                '    ClearAllControls(pnl, Recurse)
                'End If
                If ctrl.GetType() Is GetType(GroupBox) Then
                    Dim grbx As GroupBox = TryCast(ctrl, GroupBox)
                    ClearAllControls(grbx, Recurse)
                End If
            End If
        Next
    End Sub
    Public Function NoErrors(Optional Report As Boolean = False) As Boolean
        If Not String.IsNullOrEmpty(Exception) Then
            If Report = True Then MsgBox(Exception)
            Return False
        Else
            Return True
        End If
    End Function
End Class
