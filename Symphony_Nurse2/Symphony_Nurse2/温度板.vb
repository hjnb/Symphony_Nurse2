Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Public Class 温度板

    Private Sub 温度板_Click(sender As Object, e As System.EventArgs) Handles Me.Click
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        lblName.Text = TopForm.userNam.Text
        lblID.Text = TopForm.userId.Text

        With DataGridView1
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .RowHeadersVisible = False '行ヘッダー削除
            .ReadOnly = True '編集禁止
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .RowTemplate.Height = 18
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
        End With

        SQLCm.CommandText = "SELECT autono, Id, Ymd AS 日付, Hm AS 時刻, Tanto As 記載者, Ondo AS 体温, Myaku As 脈, AtuU As 血圧（上）, AtuL As 血圧（下）, Ketu As 血糖値, Weight As 体重, Height As 身長, Syoti As 処置, Val As 補記 FROM OndoD WHERE Id = " & Val(lblID.Text) & " ORDER BY Ymd, Hm"
        Adapter.Fill(Table)
        DataGridView1.DataSource = Table


        '表示設定
        With DataGridView1
            '文字を中央に表示
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            '幅
            .Columns(2).Width = 70
            .Columns(3).Width = 40
            .Columns(4).Width = 60
            .Columns(5).Width = 60
            .Columns(6).Width = 40
            .Columns(7).Width = 75
            .Columns(8).Width = 75
            .Columns(9).Width = 60
            .Columns(10).Width = 60
            .Columns(11).Width = 60
            .Columns(12).Width = 200
            .Columns(13).Width = 70

            '非表示
            .Columns("autono").Visible = False
            .Columns("Id").Visible = False
        End With

        'スクロールバーを一番下に持っていく
        DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.Rows.Count - 1

        KeyPreview = True

        Dim SQLCm2 As OleDbCommand = Cn.CreateCommand
        Dim Adapter2 As New OleDbDataAdapter(SQLCm2)
        Dim Table2 As New DataTable

        SQLCm2.CommandText = "SELECT * FROM SyotiM ORDER BY Dsp, Bunrui1, Bunrui2, Bunrui3 DESC"
        Adapter2.Fill(Table2)
        DataGridView2.DataSource = Table2
    End Sub

    Private Sub 温度板_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim forward As Boolean = e.Modifiers <> Keys.Shift
            Me.SelectNextControl(Me.ActiveControl, forward, True, True, True)
            e.Handled = True
        End If
    End Sub

    Private Sub 温度板_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
        Me.DesktopLocation = New Point(0, 50)

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        lblName.Text = TopForm.userNam.Text
        lblID.Text = TopForm.userId.Text

        Dim dtNow As DateTime = DateTime.Now
        Dim tsNow As TimeSpan = dtNow.TimeOfDay
        Dim b As String = tsNow.Hours
        Dim c As String = tsNow.Minutes
        Dim lenB As Integer = b.Length
        If lenB = 1 Then
            TimeBox1.HourText = "0" + b
        Else
            TimeBox1.HourText = b
        End If
        Dim lenC As Integer = c.Length
        If lenC = 1 Then
            TimeBox1.MinuteText = "0" + c
        Else
            TimeBox1.MinuteText = c
        End If

        With DataGridView1
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .RowHeadersVisible = False '行ヘッダー削除
            .ReadOnly = True '編集禁止
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .RowTemplate.Height = 18
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
        End With

        SQLCm.CommandText = "SELECT autono, Id, Ymd AS 日付, Hm AS 時刻, Tanto As 記載者, Ondo AS 体温, Myaku As 脈, AtuU As 血圧（上）, AtuL As 血圧（下）, Ketu As 血糖値, Weight As 体重, Height As 身長, Syoti As 処置, Val As 補記 FROM OndoD WHERE Id = " & 0 & " ORDER BY Ymd, Hm"
        Adapter.Fill(Table)
        DataGridView1.DataSource = Table


        '表示設定
        With DataGridView1
            '文字を中央に表示
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            '幅
            .Columns(2).Width = 70
            .Columns(3).Width = 40
            .Columns(4).Width = 60
            .Columns(5).Width = 60
            .Columns(6).Width = 40
            .Columns(7).Width = 75
            .Columns(8).Width = 75
            .Columns(9).Width = 60
            .Columns(10).Width = 60
            .Columns(11).Width = 60
            .Columns(12).Width = 200
            .Columns(13).Width = 70

            '非表示
            .Columns("autono").Visible = False
            .Columns("Id").Visible = False
        End With

        'スクロールバーを一番下に持っていく
        DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.Rows.Count - 1

        KeyPreview = True

        Dim SQLCm2 As OleDbCommand = Cn.CreateCommand
        Dim Adapter2 As New OleDbDataAdapter(SQLCm2)
        Dim Table2 As New DataTable

        SQLCm2.CommandText = "SELECT * FROM SyotiM ORDER BY Dsp, Bunrui1, Bunrui2, Bunrui3 DESC"
        Adapter2.Fill(Table2)
        DataGridView2.DataSource = Table2

        Dim SQLCm3 As OleDbCommand = Cn.CreateCommand
        Dim Adapter3 As New OleDbDataAdapter(SQLCm3)
        Dim Table3 As New DataTable

        SQLCm3.CommandText = "SELECT * FROM EtcM WHERE Komoku = '記載者' ORDER BY autono"
        Adapter3.Fill(Table3)
        DataGridView3.DataSource = Table3

        For i As Integer = 0 To DataGridView3.Rows.Count - 2
            cmbKisaisya.Items.Add(DataGridView3(2, i).Value)
        Next

    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        Dim r As Integer = DataGridView1.CurrentRow.Index
        YmdBox1.setADStr(DataGridView1(2, r).Value)
        TimeBox1.HourText = Strings.Left(DataGridView1(3, r).Value, 2)
        TimeBox1.MinuteText = Strings.Right(DataGridView1(3, r).Value, 2)
        cmbKisaisya.Text = DataGridView1(4, r).Value
        txtTaionn.Text = DataGridView1(5, r).Value
        txtMyaku.Text = DataGridView1(6, r).Value
        txtKetuatuUe.Text = DataGridView1(7, r).Value
        txtKetuatuSita.Text = DataGridView1(8, r).Value
        txtKettouti.Text = DataGridView1(9, r).Value
        If IsDBNull(DataGridView1(10, r).Value) = True Then

        Else
            txtSinntyou.Text = DataGridView1(10, r).Value
        End If
        txtTaijuu.Text = DataGridView1(11, r).Value
        txtSyoti1.Text = DataGridView1(12, r).Value
        txtSyoti2.Text = DataGridView1(13, r).Value
    End Sub

    Private Sub cmbSyoti_TextChanged(sender As Object, e As System.EventArgs) Handles cmbSyoti.TextChanged
        If cmbSyoti.Text = "治療プラン" Then
            lstSyoti.Items.Clear()
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                If DataGridView2(1, i).Value = "治療プラン" Then
                    If DataGridView2(4, i).Value = "" Then
                        lstSyoti.Items.Add(DataGridView2(3, i).Value)
                    ElseIf DataGridView2(4, i).Value <> "" Then
                        lstSyoti.Items.Add(DataGridView2(3, i).Value & " | " & DataGridView2(4, i).Value)
                    End If
                End If
            Next
        ElseIf cmbSyoti.Text = "処置" Then
            lstSyoti.Items.Clear()
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                If DataGridView2(1, i).Value = "処置" Then
                    lstSyoti.Items.Add(DataGridView2(3, i).Value)
                End If
            Next
        ElseIf cmbSyoti.Text = "薬剤内服" Then
            lstSyoti.Items.Clear()
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                If DataGridView2(1, i).Value = "薬剤内服" Then
                    lstSyoti.Items.Add(DataGridView2(3, i).Value)
                End If
            Next
        Else
            lstSyoti.Items.Clear()
        End If
    End Sub

    Private Sub lstSyoti_Click(sender As Object, e As System.EventArgs) Handles lstSyoti.Click
        If 0 <= lstSyoti.Text.IndexOf(" | ") Then
            txtSyoti1.Text = txtSyoti1.Text & Strings.Left(lstSyoti.Text, lstSyoti.Text.IndexOf(" | "))
            txtSyoti2.Text = Strings.Right(lstSyoti.Text, 2)
        Else
            txtSyoti1.Text = txtSyoti1.Text & lstSyoti.Text
        End If
    End Sub

    Private Sub btnTouroku_Click(sender As System.Object, e As System.EventArgs) Handles btnTouroku.Click
        If YmdBox1.getADStr() = "" Then
            MsgBox("日付を入力してください。")
            Return
        End If

        Dim count As Integer = DataGridView1.Rows.Count
        For i As Integer = 0 To count - 1
            If YmdBox1.getADStr() = DataGridView1("日付", i).Value AndAlso TimeBox1.GetTime() = DataGridView1("時刻", i).Value Then
                '日付と時刻が同じ行があったら変更
                Hennkou()
                btnKousinn.PerformClick()
                Exit Sub
            End If
        Next

        '日付と時刻が同じ行がなかったら
        Tuika()
        btnKousinn.PerformClick()
    End Sub
    Private Sub Hennkou()
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim SQL As String = ""
        SQL = "DELETE FROM OndoD WHERE (Id = " & lblID.Text & ") AND (Ymd ='" & YmdBox1.getADStr() & "') AND (Hm = '" & TimeBox1.GetTime() & "')"
        SQLCm.CommandText = SQL
        Cn.Open()
        SQLCm.ExecuteNonQuery()
        Cn.Close()

        SQLCm.Dispose()
        Cn.Dispose()

        Tuika()

    End Sub
    Private Sub Tuika()
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Table As DataTable = DirectCast(DataGridView1.DataSource, DataTable)
        Dim Row As DataRow = Table.NewRow

        Row("Id") = lblID.Text
        Row("日付") = YmdBox1.getADStr()
        Row("時刻") = TimeBox1.GetTime()
        Row("記載者") = cmbKisaisya.Text
        Row("体温") = txtTaionn.Text
        Row("脈") = txtMyaku.Text
        Row("血圧（上）") = txtKetuatuUe.Text
        Row("血圧（下）") = txtKetuatuSita.Text
        Row("処置") = txtSyoti1.Text
        Row("補記") = txtSyoti2.Text
        Row("血糖値") = txtKettouti.Text
        Row("体重") = txtTaijuu.Text
        Row("身長") = txtSinntyou.Text

        Cn.Open()
        Dim SQL As String = ""
        SQL = "INSERT INTO OndoD (Id, Ymd, Hm, Tanto, Ondo, Myaku, AtuU, AtuL, Syoti, Val, Ketu, Weight, Height) VALUES ("
        SQL &= Row("Id") & ", "
        SQL &= "'" & Row("日付") & "', "
        SQL &= "'" & Row("時刻") & "', "
        SQL &= "'" & Row("記載者") & "', "
        SQL &= "'" & Row("体温") & "', "
        SQL &= "'" & Row("脈") & "', "
        SQL &= "'" & Row("血圧（上）") & "', "
        SQL &= "'" & Row("血圧（下）") & "', "
        SQL &= "'" & Row("処置") & "', "
        SQL &= "'" & Row("補記") & "', "
        SQL &= "'" & Row("血糖値") & "', "
        SQL &= "'" & Row("体重") & "', "
        SQL &= "'" & Row("身長") & "' "
        SQL &= ")"

        SQLCm.CommandText = SQL
        SQLCm.ExecuteNonQuery()
        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()
    End Sub
    Private Sub btnSakujo_Click(sender As System.Object, e As System.EventArgs) Handles btnSakujo.Click
        Dim selectedRowIndex As Integer = If(IsNothing(DataGridView1.CurrentRow), -1, DataGridView1.CurrentRow.Index)
        '選択されていない場合は処理終了
        If selectedRowIndex = -1 Then
            MsgBox("削除対象の行が選択されていません。")
            Return
        End If

        If YmdBox1.getADStr() <> DataGridView1("日付", selectedRowIndex).Value OrElse TimeBox1.GetTime() <> DataGridView1("時刻", selectedRowIndex).Value AndAlso MessageBox.Show("データを削除しますか？", "削除確認", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
            Return
        Else
            Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
            Dim SQLCm As OleDbCommand = Cn.CreateCommand
            Dim SQL As String = ""
            SQL = "DELETE FROM OndoD WHERE (Id = " & lblID.Text & ") AND (Ymd ='" & DataGridView1("日付", selectedRowIndex).Value & "') AND (Hm = '" & DataGridView1("時刻", selectedRowIndex).Value & "')"
            SQLCm.CommandText = SQL
            Cn.Open()
            SQLCm.ExecuteNonQuery()
            Cn.Close()

            SQLCm.Dispose()
            Cn.Dispose()

            MsgBox("削除しました")
            btnKousinn.PerformClick()
        End If
    End Sub

    Private Sub btnInnsatu_Click(sender As System.Object, e As System.EventArgs) Handles btnInnsatu.Click

    End Sub

    Private Sub btnKousinn_Click(sender As System.Object, e As System.EventArgs) Handles btnKousinn.Click
        btnKuria.PerformClick()

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        lblName.Text = TopForm.userNam.Text
        lblID.Text = TopForm.userId.Text

        With DataGridView1
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .RowHeadersVisible = False '行ヘッダー削除
            .ReadOnly = True '編集禁止
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .RowTemplate.Height = 18
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
        End With

        SQLCm.CommandText = "SELECT autono, Id, Ymd AS 日付, Hm AS 時刻, Tanto As 記載者, Ondo AS 体温, Myaku As 脈, AtuU As 血圧（上）, AtuL As 血圧（下）, Ketu As 血糖値, Weight As 体重, Height As 身長, Syoti As 処置, Val As 補記 FROM OndoD WHERE Id = " & Val(lblID.Text) & " ORDER BY Ymd, Hm"
        Adapter.Fill(Table)
        DataGridView1.DataSource = Table


        '表示設定
        With DataGridView1
            '文字を中央に表示
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            '幅
            .Columns(2).Width = 70
            .Columns(3).Width = 40
            .Columns(4).Width = 60
            .Columns(5).Width = 60
            .Columns(6).Width = 40
            .Columns(7).Width = 75
            .Columns(8).Width = 75
            .Columns(9).Width = 60
            .Columns(10).Width = 60
            .Columns(11).Width = 60
            .Columns(12).Width = 200
            .Columns(13).Width = 70

            '非表示
            .Columns("autono").Visible = False
            .Columns("Id").Visible = False
        End With

        'スクロールバーを一番下に持っていく
        DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.Rows.Count - 1

        KeyPreview = True
    End Sub

    Private Sub btnKuria_Click(sender As System.Object, e As System.EventArgs) Handles btnKuria.Click
        YmdBox1.clearText()
        cmbKisaisya.Text = ""
        txtTaionn.Text = ""
        txtMyaku.Text = ""
        txtKetuatuUe.Text = ""
        txtKetuatuSita.Text = ""
        txtKettouti.Text = ""
        txtSinntyou.Text = ""
        txtTaijuu.Text = ""
        cmbSyoti.Text = ""
        txtSyoti1.Text = ""
        txtSyoti2.Text = ""
    End Sub
End Class