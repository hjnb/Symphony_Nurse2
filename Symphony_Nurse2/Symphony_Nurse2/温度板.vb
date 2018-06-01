Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Public Class 温度板
    Dim d As Boolean = False
    Private Sub 温度板_Click(sender As Object, e As System.EventArgs) Handles Me.Click
        If lblID.Text = TopForm.userId.Text AndAlso lblName.Text = TopForm.userNam.Text Then

        Else
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

            If DataGridView1.Rows.Count > 0 Then
                'スクロールバーを一番下に持っていく
                DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.Rows.Count - 1
            End If

            KeyPreview = True
        End If

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
        Me.DesktopLocation = New Point(210, 55)
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        lblName.Text = TopForm.userNam.Text
        If TopForm.userId.Text = "" Then
            lblID.Text = "0"
        Else
            lblID.Text = TopForm.userId.Text
        End If


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

        SQLCm.CommandText = "SELECT autono, Id, Ymd AS 日付, Hm AS 時刻, Tanto As 記載者, Ondo AS 体温, Myaku As 脈, AtuU As 血圧（上）, AtuL As 血圧（下）, Ketu As 血糖値, Weight As 体重, Height As 身長, Syoti As 処置, Val As 補記 FROM OndoD WHERE Id = " & lblID.Text & " ORDER BY Ymd, Hm"
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

        If DataGridView1.Rows.Count > 0 Then
            'スクロールバーを一番下に持っていく
            DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.Rows.Count - 1
        End If

        KeyPreview = True
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Columns(e.ColumnIndex).Name = "日付" Then
            If e.RowIndex > 0 AndAlso NullCheck(DataGridView1(e.ColumnIndex, e.RowIndex - 1).Value) = NullCheck(e.Value) Then
                e.Value = ""
                e.FormattingApplied = True
            End If
        End If
    End Sub

    Private Function NullCheck(cellvalue As Object) As String
        Return If(IsDBNull(cellvalue), "", cellvalue)
    End Function

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        Dim r As Integer = DataGridView1.CurrentRow.Index

        AdBox1.setADStr(NullCheck(DataGridView1(2, r).Value))
        TimeBox1.HourText = Strings.Left(NullCheck(DataGridView1(3, r).Value), 2)
        TimeBox1.MinuteText = Strings.Right(NullCheck(DataGridView1(3, r).Value), 2)
        cmbKisaisya.Text = NullCheck(DataGridView1(4, r).Value)
        txtTaionn.Text = NullCheck(DataGridView1(5, r).Value)
        txtMyaku.Text = NullCheck(DataGridView1(6, r).Value)
        txtKetuatuUe.Text = NullCheck(DataGridView1(7, r).Value)
        txtKetuatuSita.Text = NullCheck(DataGridView1(8, r).Value)
        txtKettouti.Text = NullCheck(DataGridView1(9, r).Value)
        txtTaijuu.Text = NullCheck(DataGridView1(10, r).Value)
        txtSinntyou.Text = NullCheck(DataGridView1(11, r).Value)
        txtSyoti1.Text = NullCheck(DataGridView1(12, r).Value)
        txtSyoti2.Text = NullCheck(DataGridView1(13, r).Value)
    End Sub

    Private Sub cmbKisaisyaItemAdd(sender As Object, e As System.EventArgs) Handles cmbKisaisya.GotFocus
        cmbKisaisya.Items.Clear()
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
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

    Private Sub cmbSyotiItemAdd(sender As Object, e As System.EventArgs) Handles cmbSyoti.GotFocus
        'クリア
        cmbSyoti.Items.Clear()

        'コンボボックスのリスト取得、設定
        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        SQLCm.CommandText = "select distinct Bunrui1 from SyotiM order by Bunrui1"
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        While reader.Read() = True
            cmbSyoti.Items.Add(reader("Bunrui1"))
        End While
        reader.Close()
        Cn.Close()

        Dim SQLCm2 As OleDbCommand = Cn.CreateCommand
        Dim Adapter2 As New OleDbDataAdapter(SQLCm2)
        Dim Table2 As New DataTable

        SQLCm2.CommandText = "SELECT * FROM SyotiM ORDER BY Dsp, Bunrui1, Bunrui2, Bunrui3 DESC"
        Adapter2.Fill(Table2)
        DataGridView2.DataSource = Table2

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
        ElseIf cmbSyoti.Text = "" Then
            lstSyoti.Items.Clear()
        Else
            lstSyoti.Items.Clear()
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                If DataGridView2(1, i).Value = cmbSyoti.Text Then
                    lstSyoti.Items.Add(DataGridView2(3, i).Value)
                End If
            Next
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
        If AdBox1.getADStr() = "" Then
            MsgBox("日付を入力してください。")
            Return
        End If

        Dim count As Integer = DataGridView1.Rows.Count
        For i As Integer = 0 To count - 1
            If AdBox1.getADStr() = DataGridView1("日付", i).Value AndAlso TimeBox1.GetTime() = DataGridView1("時刻", i).Value Then
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
        SQL = "DELETE FROM OndoD WHERE (Id = " & lblID.Text & ") AND (Ymd ='" & AdBox1.getADStr() & "') AND (Hm = '" & TimeBox1.GetTime() & "')"
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
        Row("日付") = AdBox1.getADStr()
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

        If (AdBox1.getADStr() <> DataGridView1("日付", selectedRowIndex).Value OrElse TimeBox1.GetTime() <> DataGridView1("時刻", selectedRowIndex).Value) AndAlso MessageBox.Show("選択行のデータを削除しますか？", "削除確認", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
            Return
        Else
            If MsgBox("削除してよろしいですか？", MsgBoxStyle.YesNo + vbExclamation, "削除確認") = MsgBoxResult.Yes Then
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

                btnKousinn.PerformClick()
            End If
        End If
    End Sub

    Private Sub btnInnsatu_Click(sender As System.Object, e As System.EventArgs) Handles btnInnsatu.Click
        Dim passForm As Form = New passwordForm(TopForm.iniFilePath, 3)
        If passForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        'Page(countrowDGV1 \ 48 + 1)
        'MsgBox(countrowDGV1)

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        Dim year As String = Strings.Left(Today, 4)
        Dim twoyearago As Integer = Val(year) - 2
        Dim oneyearago As Integer = Val(year) - 1
        SQLCm.CommandText = "SELECT autono, Id, Ymd AS 日付, Hm AS 時刻, Tanto As 記載者, Ondo AS 体温, Myaku As 脈, AtuU As 血圧（上）, AtuL As 血圧（下）, Ketu As 血糖値, Weight As 体重, Height As 身長, Syoti As 処置, Val As 補記 FROM OndoD WHERE Id = " & Val(lblID.Text) & " AND (Ymd LIKE '%" & twoyearago & "%' OR Ymd LIKE '%" & oneyearago & "%' OR Ymd LIKE '%" & year & "%') ORDER BY Ymd, Hm"
        Adapter.Fill(Table)
        DataGridView4.DataSource = Table
        
        Dim countrowDGV4 As Integer = DataGridView4.Rows.Count
        If countrowDGV4 < 48 Then
            Page1()
        ElseIf countrowDGV4 < 96 Then
            Page2()
        ElseIf countrowDGV4 < 144 Then
            Page3()
        ElseIf countrowDGV4 < 192 Then
            Page4()
        ElseIf countrowDGV4 < 240 Then
            Page5()
        ElseIf countrowDGV4 < 288 Then
            Page6()
        ElseIf countrowDGV4 < 336 Then
            Page7()
        ElseIf countrowDGV4 < 384 Then
            Page8()
        ElseIf countrowDGV4 < 432 Then
            Page9()
        ElseIf countrowDGV4 < 480 Then
            Page10()
        End If

        Dim Graf As Integer = MsgBox("グラフを印刷しますか？", vbYesNo + MsgBoxStyle.DefaultButton2, "印刷")
        If Graf = 6 Then
            PrintGraf()
        End If

    End Sub
    Private Sub Page(c As Integer)
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("温度板 (" & c & ")")

        For i As Integer = 2 To 55 * c + 2 Step 55
            oSheet.Range("C" & i).Value = lblID.Text
            oSheet.Range("E" & i).Value = lblName.Text
        Next

        Dim countrowDGV1 As Integer = DataGridView1.Rows.Count
        Dim lastcount As Integer = 0
        If c = 1 Then
            lastcount = countrowDGV1
        ElseIf c = 2 Then

        ElseIf c = 3 Then

        End If


        '1枚目 c=1
        For i As Integer = 0 To 47
            oSheet.Range("B" & i + 5 + (c - 1) * 7).Value = Strings.Right(DataGridView1(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(3, i).Value)
            oSheet.Range("D" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(4, i).Value)
            oSheet.Range("E" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(5, i).Value)
            oSheet.Range("F" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(6, i).Value)
            oSheet.Range("G" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(7, i).Value)
            oSheet.Range("H" & i + 5 + (c - 1) * 7).Value = "/"
            oSheet.Range("I" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(8, i).Value)
            oSheet.Range("J" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(9, i).Value)
            oSheet.Range("K" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(10, i).Value)
            oSheet.Range("L" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(11, i).Value)
            oSheet.Range("M" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(12, i).Value)
            oSheet.Range("O" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(13, i).Value)
        Next
        '2枚目
        For i As Integer = 48 To 95
            oSheet.Range("B" & i + 5 + (c - 1) * 7).Value = Strings.Right(DataGridView1(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(3, i).Value)
            oSheet.Range("D" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(4, i).Value)
            oSheet.Range("E" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(5, i).Value)
            oSheet.Range("F" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(6, i).Value)
            oSheet.Range("G" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(7, i).Value)
            oSheet.Range("H" & i + 5 + (c - 1) * 7).Value = "/"
            oSheet.Range("I" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(8, i).Value)
            oSheet.Range("J" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(9, i).Value)
            oSheet.Range("K" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(10, i).Value)
            oSheet.Range("L" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(11, i).Value)
            oSheet.Range("M" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(12, i).Value)
            oSheet.Range("O" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(13, i).Value)
        Next
        '3枚目
        For i As Integer = 96 To countrowDGV1 - 1
            oSheet.Range("B" & i + 5 + (c - 1) * 7).Value = Strings.Right(DataGridView1(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(3, i).Value)
            oSheet.Range("D" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(4, i).Value)
            oSheet.Range("E" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(5, i).Value)
            oSheet.Range("F" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(6, i).Value)
            oSheet.Range("G" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(7, i).Value)
            oSheet.Range("H" & i + 5 + (c - 1) * 7).Value = "/"
            oSheet.Range("I" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(8, i).Value)
            oSheet.Range("J" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(9, i).Value)
            oSheet.Range("K" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(10, i).Value)
            oSheet.Range("L" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(11, i).Value)
            oSheet.Range("M" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(12, i).Value)
            oSheet.Range("O" & i + 5 + (c - 1) * 7).Value = NullCheck(DataGridView1(13, i).Value)
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        '印刷
        oSheet.PrintPreview(1)

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
    Private Sub Page1()
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object
        Dim day As DateTime = DateTime.Today

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("温度板")

        oSheet.Range("C2").Value = lblID.Text
        oSheet.Range("E2").Value = lblName.Text

        Dim countrowDGV1 As Integer = DataGridView4.Rows.Count
        '1枚目
        For i As Integer = 0 To countrowDGV1 - 1
            oSheet.Range("B" & i + 5).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 5).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 5).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 5).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 5).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 5).Value = "/"
            oSheet.Range("I" & i + 5).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 5).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 5).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 5).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 5).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 5).Value = NullCheck(DataGridView4(13, i).Value)
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        '印刷
        If TopForm.rbtnPreview.Checked = True Then
            oSheet.PrintPreview(1)
        ElseIf TopForm.rbtnPrint.Checked = True Then
            oSheet.Printout(1)
        End If


        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
    Private Sub Page2()
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object
        Dim day As DateTime = DateTime.Today

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("温度板 (2)")

        For i As Integer = 2 To 57 Step 55
            oSheet.Range("C" & i).Value = lblID.Text
            oSheet.Range("E" & i).Value = lblName.Text
        Next


        Dim countrowDGV1 As Integer = DataGridView4.Rows.Count
        '1枚目
        For i As Integer = 0 To 47
            oSheet.Range("B" & i + 5).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 5).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 5).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 5).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 5).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 5).Value = "/"
            oSheet.Range("I" & i + 5).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 5).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 5).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 5).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 5).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 5).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '2枚目
        For i As Integer = 48 To countrowDGV1 - 1
            oSheet.Range("B" & i + 12).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 12).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 12).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 12).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 12).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 12).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 12).Value = "/"
            oSheet.Range("I" & i + 12).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 12).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 12).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 12).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 12).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 12).Value = NullCheck(DataGridView4(13, i).Value)
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        If TopForm.rbtnPreview.Checked = True Then
            oSheet.PrintPreview(1)
        ElseIf TopForm.rbtnPrint.Checked = True Then
            oSheet.Printout(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
    Private Sub Page3()
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object
        Dim day As DateTime = DateTime.Today

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("温度板 (3)")

        For i As Integer = 2 To 112 Step 55
            oSheet.Range("C" & i).Value = lblID.Text
            oSheet.Range("E" & i).Value = lblName.Text
        Next


        Dim countrowDGV1 As Integer = DataGridView4.Rows.Count
        '1枚目
        For i As Integer = 0 To 47
            oSheet.Range("B" & i + 5).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 5).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 5).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 5).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 5).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 5).Value = "/"
            oSheet.Range("I" & i + 5).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 5).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 5).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 5).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 5).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 5).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '2枚目
        For i As Integer = 48 To 95
            oSheet.Range("B" & i + 12).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 12).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 12).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 12).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 12).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 12).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 12).Value = "/"
            oSheet.Range("I" & i + 12).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 12).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 12).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 12).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 12).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 12).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '3枚目
        For i As Integer = 96 To countrowDGV1 - 1
            oSheet.Range("B" & i + 19).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 19).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 19).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 19).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 19).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 19).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 19).Value = "/"
            oSheet.Range("I" & i + 19).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 19).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 19).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 19).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 19).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 19).Value = NullCheck(DataGridView4(13, i).Value)
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        '印刷
        If TopForm.rbtnPreview.Checked = True Then
            oSheet.PrintPreview(1)
        ElseIf TopForm.rbtnPrint.Checked = True Then
            oSheet.Printout(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
    Private Sub Page4()
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object
        Dim day As DateTime = DateTime.Today

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("温度板 (4)")

        For i As Integer = 2 To 167 Step 55
            oSheet.Range("C" & i).Value = lblID.Text
            oSheet.Range("E" & i).Value = lblName.Text
        Next


        Dim countrowDGV1 As Integer = DataGridView4.Rows.Count
        '1枚目
        For i As Integer = 0 To 47
            oSheet.Range("B" & i + 5).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 5).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 5).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 5).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 5).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 5).Value = "/"
            oSheet.Range("I" & i + 5).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 5).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 5).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 5).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 5).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 5).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '2枚目
        For i As Integer = 48 To 95
            oSheet.Range("B" & i + 12).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 12).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 12).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 12).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 12).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 12).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 12).Value = "/"
            oSheet.Range("I" & i + 12).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 12).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 12).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 12).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 12).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 12).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '3枚目
        For i As Integer = 96 To 143
            oSheet.Range("B" & i + 19).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 19).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 19).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 19).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 19).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 19).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 19).Value = "/"
            oSheet.Range("I" & i + 19).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 19).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 19).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 19).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 19).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 19).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '4枚目
        For i As Integer = 144 To countrowDGV1 - 1
            oSheet.Range("B" & i + 26).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 26).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 26).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 26).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 26).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 26).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 26).Value = "/"
            oSheet.Range("I" & i + 26).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 26).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 26).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 26).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 26).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 26).Value = NullCheck(DataGridView4(13, i).Value)
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        '印刷
        If TopForm.rbtnPreview.Checked = True Then
            oSheet.PrintPreview(1)
        ElseIf TopForm.rbtnPrint.Checked = True Then
            oSheet.Printout(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
    Private Sub Page5()
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object
        Dim day As DateTime = DateTime.Today

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("温度板 (5)")

        For i As Integer = 2 To 222 Step 55
            oSheet.Range("C" & i).Value = lblID.Text
            oSheet.Range("E" & i).Value = lblName.Text
        Next


        Dim countrowDGV1 As Integer = DataGridView4.Rows.Count
        '1枚目
        For i As Integer = 0 To 47
            oSheet.Range("B" & i + 5).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 5).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 5).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 5).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 5).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 5).Value = "/"
            oSheet.Range("I" & i + 5).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 5).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 5).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 5).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 5).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 5).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '2枚目
        For i As Integer = 48 To 95
            oSheet.Range("B" & i + 12).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 12).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 12).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 12).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 12).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 12).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 12).Value = "/"
            oSheet.Range("I" & i + 12).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 12).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 12).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 12).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 12).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 12).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '3枚目
        For i As Integer = 96 To 143
            oSheet.Range("B" & i + 19).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 19).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 19).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 19).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 19).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 19).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 19).Value = "/"
            oSheet.Range("I" & i + 19).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 19).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 19).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 19).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 19).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 19).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '4枚目
        For i As Integer = 144 To 191
            oSheet.Range("B" & i + 26).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 26).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 26).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 26).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 26).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 26).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 26).Value = "/"
            oSheet.Range("I" & i + 26).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 26).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 26).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 26).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 26).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 26).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '5枚目
        For i As Integer = 192 To countrowDGV1 - 1
            oSheet.Range("B" & i + 33).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 33).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 33).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 33).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 33).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 33).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 33).Value = "/"
            oSheet.Range("I" & i + 33).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 33).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 33).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 33).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 33).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 33).Value = NullCheck(DataGridView4(13, i).Value)
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        '印刷
        If TopForm.rbtnPreview.Checked = True Then
            oSheet.PrintPreview(1)
        ElseIf TopForm.rbtnPrint.Checked = True Then
            oSheet.Printout(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
    Private Sub Page6()
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object
        Dim day As DateTime = DateTime.Today

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("温度板 (6)")

        For i As Integer = 2 To 277 Step 55
            oSheet.Range("C" & i).Value = lblID.Text
            oSheet.Range("E" & i).Value = lblName.Text
        Next


        Dim countrowDGV1 As Integer = DataGridView4.Rows.Count
        '1枚目
        For i As Integer = 0 To 47
            oSheet.Range("B" & i + 5).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 5).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 5).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 5).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 5).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 5).Value = "/"
            oSheet.Range("I" & i + 5).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 5).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 5).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 5).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 5).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 5).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '2枚目
        For i As Integer = 48 To 95
            oSheet.Range("B" & i + 12).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 12).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 12).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 12).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 12).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 12).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 12).Value = "/"
            oSheet.Range("I" & i + 12).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 12).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 12).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 12).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 12).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 12).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '3枚目
        For i As Integer = 96 To 143
            oSheet.Range("B" & i + 19).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 19).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 19).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 19).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 19).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 19).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 19).Value = "/"
            oSheet.Range("I" & i + 19).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 19).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 19).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 19).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 19).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 19).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '4枚目
        For i As Integer = 144 To 191
            oSheet.Range("B" & i + 26).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 26).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 26).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 26).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 26).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 26).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 26).Value = "/"
            oSheet.Range("I" & i + 26).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 26).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 26).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 26).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 26).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 26).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '5枚目
        For i As Integer = 192 To 239
            oSheet.Range("B" & i + 33).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 33).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 33).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 33).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 33).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 33).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 33).Value = "/"
            oSheet.Range("I" & i + 33).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 33).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 33).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 33).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 33).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 33).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '6枚目
        For i As Integer = 240 To countrowDGV1 - 1
            oSheet.Range("B" & i + 40).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 40).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 40).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 40).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 40).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 40).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 40).Value = "/"
            oSheet.Range("I" & i + 40).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 40).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 40).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 40).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 40).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 40).Value = NullCheck(DataGridView4(13, i).Value)
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        '印刷
        If TopForm.rbtnPreview.Checked = True Then
            oSheet.PrintPreview(1)
        ElseIf TopForm.rbtnPrint.Checked = True Then
            oSheet.Printout(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
    Private Sub Page7()
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object
        Dim day As DateTime = DateTime.Today

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("温度板 (7)")

        For i As Integer = 2 To 332 Step 55
            oSheet.Range("C" & i).Value = lblID.Text
            oSheet.Range("E" & i).Value = lblName.Text
        Next


        Dim countrowDGV1 As Integer = DataGridView4.Rows.Count
        '1枚目
        For i As Integer = 0 To 47
            oSheet.Range("B" & i + 5).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 5).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 5).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 5).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 5).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 5).Value = "/"
            oSheet.Range("I" & i + 5).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 5).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 5).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 5).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 5).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 5).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '2枚目
        For i As Integer = 48 To 95
            oSheet.Range("B" & i + 12).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 12).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 12).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 12).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 12).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 12).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 12).Value = "/"
            oSheet.Range("I" & i + 12).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 12).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 12).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 12).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 12).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 12).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '3枚目
        For i As Integer = 96 To 143
            oSheet.Range("B" & i + 19).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 19).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 19).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 19).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 19).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 19).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 19).Value = "/"
            oSheet.Range("I" & i + 19).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 19).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 19).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 19).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 19).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 19).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '4枚目
        For i As Integer = 144 To 191
            oSheet.Range("B" & i + 26).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 26).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 26).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 26).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 26).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 26).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 26).Value = "/"
            oSheet.Range("I" & i + 26).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 26).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 26).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 26).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 26).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 26).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '5枚目
        For i As Integer = 192 To 239
            oSheet.Range("B" & i + 33).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 33).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 33).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 33).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 33).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 33).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 33).Value = "/"
            oSheet.Range("I" & i + 33).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 33).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 33).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 33).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 33).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 33).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '6枚目
        For i As Integer = 240 To 287
            oSheet.Range("B" & i + 40).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 40).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 40).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 40).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 40).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 40).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 40).Value = "/"
            oSheet.Range("I" & i + 40).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 40).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 40).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 40).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 40).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 40).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '7枚目
        For i As Integer = 288 To countrowDGV1 - 1
            oSheet.Range("B" & i + 47).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 47).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 47).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 47).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 47).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 47).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 47).Value = "/"
            oSheet.Range("I" & i + 47).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 47).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 47).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 47).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 47).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 47).Value = NullCheck(DataGridView4(13, i).Value)
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        '印刷
        If TopForm.rbtnPreview.Checked = True Then
            oSheet.PrintPreview(1)
        ElseIf TopForm.rbtnPrint.Checked = True Then
            oSheet.Printout(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
    Private Sub Page8()
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object
        Dim day As DateTime = DateTime.Today

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("温度板 (8)")

        For i As Integer = 2 To 387 Step 55
            oSheet.Range("C" & i).Value = lblID.Text
            oSheet.Range("E" & i).Value = lblName.Text
        Next


        Dim countrowDGV1 As Integer = DataGridView4.Rows.Count
        '1枚目
        For i As Integer = 0 To 47
            oSheet.Range("B" & i + 5).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 5).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 5).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 5).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 5).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 5).Value = "/"
            oSheet.Range("I" & i + 5).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 5).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 5).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 5).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 5).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 5).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '2枚目
        For i As Integer = 48 To 95
            oSheet.Range("B" & i + 12).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 12).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 12).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 12).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 12).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 12).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 12).Value = "/"
            oSheet.Range("I" & i + 12).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 12).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 12).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 12).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 12).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 12).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '3枚目
        For i As Integer = 96 To 143
            oSheet.Range("B" & i + 19).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 19).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 19).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 19).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 19).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 19).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 19).Value = "/"
            oSheet.Range("I" & i + 19).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 19).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 19).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 19).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 19).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 19).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '4枚目
        For i As Integer = 144 To 191
            oSheet.Range("B" & i + 26).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 26).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 26).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 26).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 26).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 26).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 26).Value = "/"
            oSheet.Range("I" & i + 26).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 26).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 26).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 26).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 26).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 26).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '5枚目
        For i As Integer = 192 To 239
            oSheet.Range("B" & i + 33).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 33).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 33).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 33).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 33).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 33).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 33).Value = "/"
            oSheet.Range("I" & i + 33).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 33).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 33).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 33).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 33).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 33).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '6枚目
        For i As Integer = 240 To 287
            oSheet.Range("B" & i + 40).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 40).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 40).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 40).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 40).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 40).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 40).Value = "/"
            oSheet.Range("I" & i + 40).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 40).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 40).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 40).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 40).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 40).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '7枚目
        For i As Integer = 288 To 335
            oSheet.Range("B" & i + 47).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 47).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 47).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 47).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 47).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 47).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 47).Value = "/"
            oSheet.Range("I" & i + 47).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 47).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 47).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 47).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 47).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 47).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '8枚目
        For i As Integer = 336 To countrowDGV1 - 1
            oSheet.Range("B" & i + 54).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 54).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 54).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 54).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 54).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 54).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 54).Value = "/"
            oSheet.Range("I" & i + 54).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 54).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 54).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 54).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 54).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 54).Value = NullCheck(DataGridView4(13, i).Value)
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        '印刷
        If TopForm.rbtnPreview.Checked = True Then
            oSheet.PrintPreview(1)
        ElseIf TopForm.rbtnPrint.Checked = True Then
            oSheet.Printout(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
    Private Sub Page9()
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("温度板 (9)")

        For i As Integer = 2 To 442 Step 55
            oSheet.Range("C" & i).Value = lblID.Text
            oSheet.Range("E" & i).Value = lblName.Text
        Next


        Dim countrowDGV1 As Integer = DataGridView4.Rows.Count
        '1枚目
        For i As Integer = 0 To 47
            oSheet.Range("B" & i + 5).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 5).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 5).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 5).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 5).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 5).Value = "/"
            oSheet.Range("I" & i + 5).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 5).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 5).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 5).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 5).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 5).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '2枚目
        For i As Integer = 48 To 95
            oSheet.Range("B" & i + 12).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 12).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 12).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 12).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 12).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 12).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 12).Value = "/"
            oSheet.Range("I" & i + 12).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 12).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 12).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 12).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 12).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 12).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '3枚目
        For i As Integer = 96 To 143
            oSheet.Range("B" & i + 19).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 19).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 19).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 19).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 19).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 19).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 19).Value = "/"
            oSheet.Range("I" & i + 19).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 19).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 19).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 19).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 19).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 19).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '4枚目
        For i As Integer = 144 To 191
            oSheet.Range("B" & i + 26).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 26).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 26).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 26).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 26).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 26).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 26).Value = "/"
            oSheet.Range("I" & i + 26).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 26).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 26).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 26).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 26).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 26).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '5枚目
        For i As Integer = 192 To 239
            oSheet.Range("B" & i + 33).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 33).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 33).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 33).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 33).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 33).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 33).Value = "/"
            oSheet.Range("I" & i + 33).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 33).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 33).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 33).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 33).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 33).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '6枚目
        For i As Integer = 240 To 287
            oSheet.Range("B" & i + 40).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 40).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 40).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 40).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 40).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 40).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 40).Value = "/"
            oSheet.Range("I" & i + 40).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 40).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 40).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 40).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 40).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 40).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '7枚目
        For i As Integer = 288 To 335
            oSheet.Range("B" & i + 47).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 47).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 47).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 47).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 47).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 47).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 47).Value = "/"
            oSheet.Range("I" & i + 47).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 47).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 47).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 47).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 47).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 47).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '8枚目
        For i As Integer = 336 To 383
            oSheet.Range("B" & i + 54).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 54).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 54).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 54).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 54).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 54).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 54).Value = "/"
            oSheet.Range("I" & i + 54).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 54).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 54).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 54).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 54).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 54).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '9枚目
        For i As Integer = 384 To countrowDGV1 - 1
            oSheet.Range("B" & i + 61).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 61).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 61).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 61).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 61).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 61).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 61).Value = "/"
            oSheet.Range("I" & i + 61).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 61).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 61).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 61).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 61).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 61).Value = NullCheck(DataGridView4(13, i).Value)
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        '印刷
        If TopForm.rbtnPreview.Checked = True Then
            oSheet.PrintPreview(1)
        ElseIf TopForm.rbtnPrint.Checked = True Then
            oSheet.Printout(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
    Private Sub Page10()
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("温度板 (10)")

        For i As Integer = 2 To 497 Step 55
            oSheet.Range("C" & i).Value = lblID.Text
            oSheet.Range("E" & i).Value = lblName.Text
        Next


        Dim countrowDGV1 As Integer = DataGridView4.Rows.Count
        '1枚目
        For i As Integer = 0 To 47
            oSheet.Range("B" & i + 5).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 5).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 5).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 5).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 5).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 5).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 5).Value = "/"
            oSheet.Range("I" & i + 5).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 5).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 5).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 5).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 5).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 5).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '2枚目
        For i As Integer = 48 To 95
            oSheet.Range("B" & i + 12).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 12).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 12).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 12).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 12).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 12).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 12).Value = "/"
            oSheet.Range("I" & i + 12).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 12).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 12).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 12).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 12).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 12).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '3枚目
        For i As Integer = 96 To 143
            oSheet.Range("B" & i + 19).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 19).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 19).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 19).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 19).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 19).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 19).Value = "/"
            oSheet.Range("I" & i + 19).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 19).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 19).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 19).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 19).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 19).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '4枚目
        For i As Integer = 144 To 191
            oSheet.Range("B" & i + 26).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 26).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 26).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 26).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 26).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 26).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 26).Value = "/"
            oSheet.Range("I" & i + 26).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 26).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 26).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 26).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 26).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 26).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '5枚目
        For i As Integer = 192 To 239
            oSheet.Range("B" & i + 33).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 33).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 33).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 33).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 33).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 33).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 33).Value = "/"
            oSheet.Range("I" & i + 33).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 33).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 33).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 33).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 33).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 33).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '6枚目
        For i As Integer = 240 To 287
            oSheet.Range("B" & i + 40).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 40).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 40).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 40).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 40).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 40).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 40).Value = "/"
            oSheet.Range("I" & i + 40).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 40).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 40).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 40).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 40).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 40).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '7枚目
        For i As Integer = 288 To 335
            oSheet.Range("B" & i + 47).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 47).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 47).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 47).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 47).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 47).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 47).Value = "/"
            oSheet.Range("I" & i + 47).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 47).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 47).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 47).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 47).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 47).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '8枚目
        For i As Integer = 336 To 383
            oSheet.Range("B" & i + 54).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 54).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 54).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 54).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 54).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 54).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 54).Value = "/"
            oSheet.Range("I" & i + 54).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 54).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 54).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 54).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 54).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 54).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '9枚目
        For i As Integer = 384 To 431
            oSheet.Range("B" & i + 61).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 61).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 61).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 61).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 61).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 61).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 61).Value = "/"
            oSheet.Range("I" & i + 61).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 61).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 61).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 61).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 61).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 61).Value = NullCheck(DataGridView4(13, i).Value)
        Next
        '10枚目
        For i As Integer = 432 To countrowDGV1 - 1
            oSheet.Range("B" & i + 68).Value = Strings.Right(DataGridView4(2, i).FormattedValue, 5)
            oSheet.Range("C" & i + 68).Value = NullCheck(DataGridView4(3, i).Value)
            oSheet.Range("D" & i + 68).Value = NullCheck(DataGridView4(4, i).Value)
            oSheet.Range("E" & i + 68).Value = NullCheck(DataGridView4(5, i).Value)
            oSheet.Range("F" & i + 68).Value = NullCheck(DataGridView4(6, i).Value)
            oSheet.Range("G" & i + 68).Value = NullCheck(DataGridView4(7, i).Value)
            oSheet.Range("H" & i + 68).Value = "/"
            oSheet.Range("I" & i + 68).Value = NullCheck(DataGridView4(8, i).Value)
            oSheet.Range("J" & i + 68).Value = NullCheck(DataGridView4(9, i).Value)
            oSheet.Range("K" & i + 68).Value = NullCheck(DataGridView4(10, i).Value)
            oSheet.Range("L" & i + 68).Value = NullCheck(DataGridView4(11, i).Value)
            oSheet.Range("M" & i + 68).Value = NullCheck(DataGridView4(12, i).Value)
            oSheet.Range("O" & i + 68).Value = NullCheck(DataGridView4(13, i).Value)
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        '印刷
        If TopForm.rbtnPreview.Checked = True Then
            oSheet.PrintPreview(1)
        ElseIf TopForm.rbtnPrint.Checked = True Then
            oSheet.Printout(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
    Private Sub PrintGraf()

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        Dim yearmonth As String = Strings.Left(Today, 7)
        
        SQLCm.CommandText = "SELECT autono, Id, Ymd AS 日付, Hm AS 時刻, Tanto As 記載者, Ondo AS 体温, Myaku As 脈, AtuU As 血圧（上）, AtuL As 血圧（下）, Ketu As 血糖値, Weight As 体重, Height As 身長, Syoti As 処置, Val As 補記 FROM OndoD WHERE Id = " & Val(lblID.Text) & " AND Ymd LIKE '%" & yearmonth & "%' ORDER BY Ymd, Hm"
        Adapter.Fill(Table)
        DataGridView4.DataSource = Table

        If DataGridView4.Rows.Count = 0 Then
            MsgBox("今月のデータがありません")
            Return
        End If

        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object
        Dim oSheet2 As Object
        Dim day As DateTime = DateTime.Today

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("グラフ新")
        oSheet2 = objWorkBook.Worksheets("グラフ新 (2)")

        oSheet2.Range("K2").Value = lblID.Text
        oSheet2.Range("S2").Value = lblName.Text
        oSheet2.Range("AQ2").Value = yearmonth
        oSheet2.Range("CD2").Value = "1"

        Dim countrowDGV1 As Integer = DataGridView4.Rows.Count
        '1枚目
        For i As Integer = 0 To countrowDGV1 - 1
            If Strings.Right(DataGridView4(2, i).Value, 2) = "01" Then
                oSheet.Range("CV5").Value = DataGridView4(7, i).Value
                oSheet.Range("CV6").Value = DataGridView4(8, i).Value
                oSheet.Range("CV7").Value = DataGridView4(6, i).Value
                oSheet.Range("CV8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "02" Then
                oSheet.Range("CY5").Value = DataGridView4(7, i).Value
                oSheet.Range("CY6").Value = DataGridView4(8, i).Value
                oSheet.Range("CY7").Value = DataGridView4(6, i).Value
                oSheet.Range("CY8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "03" Then
                oSheet.Range("DB5").Value = DataGridView4(7, i).Value
                oSheet.Range("DB6").Value = DataGridView4(8, i).Value
                oSheet.Range("DB7").Value = DataGridView4(6, i).Value
                oSheet.Range("DB8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "04" Then
                oSheet.Range("DE5").Value = DataGridView4(7, i).Value
                oSheet.Range("DE6").Value = DataGridView4(8, i).Value
                oSheet.Range("DE7").Value = DataGridView4(6, i).Value
                oSheet.Range("DE8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "05" Then
                oSheet.Range("DH5").Value = DataGridView4(7, i).Value
                oSheet.Range("DH6").Value = DataGridView4(8, i).Value
                oSheet.Range("DH7").Value = DataGridView4(6, i).Value
                oSheet.Range("DH8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "06" Then
                oSheet.Range("DK5").Value = DataGridView4(7, i).Value
                oSheet.Range("DK6").Value = DataGridView4(8, i).Value
                oSheet.Range("DK7").Value = DataGridView4(6, i).Value
                oSheet.Range("DK8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "07" Then
                oSheet.Range("DN5").Value = DataGridView4(7, i).Value
                oSheet.Range("DN6").Value = DataGridView4(8, i).Value
                oSheet.Range("DN7").Value = DataGridView4(6, i).Value
                oSheet.Range("DN8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "08" Then
                oSheet.Range("DQ5").Value = DataGridView4(7, i).Value
                oSheet.Range("DQ6").Value = DataGridView4(8, i).Value
                oSheet.Range("DQ7").Value = DataGridView4(6, i).Value
                oSheet.Range("DQ8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "09" Then
                oSheet.Range("DT5").Value = DataGridView4(7, i).Value
                oSheet.Range("DT6").Value = DataGridView4(8, i).Value
                oSheet.Range("DT7").Value = DataGridView4(6, i).Value
                oSheet.Range("DT8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "10" Then
                oSheet.Range("DW5").Value = DataGridView4(7, i).Value
                oSheet.Range("DW6").Value = DataGridView4(8, i).Value
                oSheet.Range("DW7").Value = DataGridView4(6, i).Value
                oSheet.Range("DW8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "11" Then
                oSheet.Range("DZ5").Value = DataGridView4(7, i).Value
                oSheet.Range("DZ6").Value = DataGridView4(8, i).Value
                oSheet.Range("DZ7").Value = DataGridView4(6, i).Value
                oSheet.Range("DZ8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "12" Then
                oSheet.Range("EC5").Value = DataGridView4(7, i).Value
                oSheet.Range("EC6").Value = DataGridView4(8, i).Value
                oSheet.Range("EC7").Value = DataGridView4(6, i).Value
                oSheet.Range("EC8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "13" Then
                oSheet.Range("EF5").Value = DataGridView4(7, i).Value
                oSheet.Range("EF6").Value = DataGridView4(8, i).Value
                oSheet.Range("EF7").Value = DataGridView4(6, i).Value
                oSheet.Range("EF8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "14" Then
                oSheet.Range("EI5").Value = DataGridView4(7, i).Value
                oSheet.Range("EI6").Value = DataGridView4(8, i).Value
                oSheet.Range("EI7").Value = DataGridView4(6, i).Value
                oSheet.Range("EI8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "15" Then
                oSheet.Range("EL5").Value = DataGridView4(7, i).Value
                oSheet.Range("EL6").Value = DataGridView4(8, i).Value
                oSheet.Range("EL7").Value = DataGridView4(6, i).Value
                oSheet.Range("EL8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "16" Then
                oSheet.Range("EO5").Value = DataGridView4(7, i).Value
                oSheet.Range("EO6").Value = DataGridView4(8, i).Value
                oSheet.Range("EO7").Value = DataGridView4(6, i).Value
                oSheet.Range("EO8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "17" Then
                oSheet.Range("ER5").Value = DataGridView4(7, i).Value
                oSheet.Range("ER6").Value = DataGridView4(8, i).Value
                oSheet.Range("ER7").Value = DataGridView4(6, i).Value
                oSheet.Range("ER8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "18" Then
                oSheet.Range("EU5").Value = DataGridView4(7, i).Value
                oSheet.Range("EU6").Value = DataGridView4(8, i).Value
                oSheet.Range("EU7").Value = DataGridView4(6, i).Value
                oSheet.Range("EU8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "19" Then
                oSheet.Range("EX5").Value = DataGridView4(7, i).Value
                oSheet.Range("EX6").Value = DataGridView4(8, i).Value
                oSheet.Range("EX7").Value = DataGridView4(6, i).Value
                oSheet.Range("EX8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "20" Then
                oSheet.Range("FA5").Value = DataGridView4(7, i).Value
                oSheet.Range("FA6").Value = DataGridView4(8, i).Value
                oSheet.Range("FA7").Value = DataGridView4(6, i).Value
                oSheet.Range("FA8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "21" Then
                oSheet.Range("FD5").Value = DataGridView4(7, i).Value
                oSheet.Range("FD6").Value = DataGridView4(8, i).Value
                oSheet.Range("FD7").Value = DataGridView4(6, i).Value
                oSheet.Range("FD8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "22" Then
                oSheet.Range("FG5").Value = DataGridView4(7, i).Value
                oSheet.Range("FG6").Value = DataGridView4(8, i).Value
                oSheet.Range("FG7").Value = DataGridView4(6, i).Value
                oSheet.Range("FG8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "23" Then
                oSheet.Range("FJ5").Value = DataGridView4(7, i).Value
                oSheet.Range("FJ6").Value = DataGridView4(8, i).Value
                oSheet.Range("FJ7").Value = DataGridView4(6, i).Value
                oSheet.Range("FJ8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "24" Then
                oSheet.Range("FM5").Value = DataGridView4(7, i).Value
                oSheet.Range("FM6").Value = DataGridView4(8, i).Value
                oSheet.Range("FM7").Value = DataGridView4(6, i).Value
                oSheet.Range("FM8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "25" Then
                oSheet.Range("FP5").Value = DataGridView4(7, i).Value
                oSheet.Range("FP6").Value = DataGridView4(8, i).Value
                oSheet.Range("FP7").Value = DataGridView4(6, i).Value
                oSheet.Range("FP8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "26" Then
                oSheet.Range("FP5").Value = DataGridView4(7, i).Value
                oSheet.Range("FP6").Value = DataGridView4(8, i).Value
                oSheet.Range("FP7").Value = DataGridView4(6, i).Value
                oSheet.Range("FP8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "27" Then
                oSheet.Range("FS5").Value = DataGridView4(7, i).Value
                oSheet.Range("FS6").Value = DataGridView4(8, i).Value
                oSheet.Range("FS7").Value = DataGridView4(6, i).Value
                oSheet.Range("FS8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "28" Then
                oSheet.Range("FV5").Value = DataGridView4(7, i).Value
                oSheet.Range("FV6").Value = DataGridView4(8, i).Value
                oSheet.Range("FV7").Value = DataGridView4(6, i).Value
                oSheet.Range("FV8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "29" Then
                oSheet.Range("FY5").Value = DataGridView4(7, i).Value
                oSheet.Range("FY6").Value = DataGridView4(8, i).Value
                oSheet.Range("FY7").Value = DataGridView4(6, i).Value
                oSheet.Range("FY8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "30" Then
                oSheet.Range("CY5").Value = DataGridView4(7, i).Value
                oSheet.Range("CY6").Value = DataGridView4(8, i).Value
                oSheet.Range("CY7").Value = DataGridView4(6, i).Value
                oSheet.Range("CY8").Value = DataGridView4(5, i).Value
            ElseIf Strings.Right(DataGridView4(2, i).Value, 2) = "31" Then
                oSheet.Range("GB5").Value = DataGridView4(7, i).Value
                oSheet.Range("GB6").Value = DataGridView4(8, i).Value
                oSheet.Range("GB7").Value = DataGridView4(6, i).Value
                oSheet.Range("GB8").Value = DataGridView4(5, i).Value
            End If
        Next

        '保存
        objExcel.DisplayAlerts = False

        ' エクセル表示
        objExcel.Visible = True

        '印刷
        If TopForm.rbtnPreview.Checked = True Then
            oSheet2.PrintPreview(1)
        ElseIf TopForm.rbtnPrint.Checked = True Then
            oSheet2.Printout(1)
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(oSheet2)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
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
        AdBox1.setADStr(Today.ToString("yyyy/MM/dd"))
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

    Private Sub btnTimeUe_Click(sender As System.Object, e As System.EventArgs)



    End Sub

    Private Sub btnTimeSita_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub textBoxinput_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtTaionn.KeyDown, txtMyaku.KeyDown, txtKetuatuUe.KeyDown, txtKetuatuSita.KeyDown, txtKettouti.KeyDown, txtSinntyou.KeyDown
        'アクティブなコントロールを取得
        Dim cControl As TextBox = Me.ActiveControl  'イベント発生は指定のテキストボックスのときだけなのでTextBox型で宣言

        If e.KeyCode = Windows.Forms.Keys.Left Then
            'カーソルを左に動かす
        ElseIf e.KeyCode = Windows.Forms.Keys.Right Then
            'カーソルを右に動かす
        ElseIf (Keys.NumPad0 <= e.KeyCode AndAlso e.KeyCode <= Keys.NumPad9) OrElse (Keys.D0 <= e.KeyCode AndAlso e.KeyCode <= Keys.D9) Then
            '数字の処理
        ElseIf e.KeyCode = Keys.Decimal OrElse e.KeyCode = 190 Then
            '小数点の処理
        ElseIf e.KeyCode = Keys.Delete Then    'Delキーが押されたとき
            'If cControl.Name = "txtTaijuu" AndAlso Val(txtTaijuu.Text) = 0 Then     'アクティブなコントロールがtxtTaijuuでテキストの中身が0のとき
            '    e.SuppressKeyPress = True   '何もしない
            'ElseIf cControl.Name = "txtTaijuu" AndAlso Val(txtTaijuu.Text) <> 0 Then     'アクティブなコントロールがtxtTaijuuでテキストの中身が0ではないとき
            '    txtTaijuu.Text = "0.00"     '値を0.00にする
            '    e.SuppressKeyPress = True
            'Else
            '    'Delの処理
            'End If
        Else
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub txtTaijuu_Click(sender As Object, e As System.EventArgs) Handles txtTaijuu.Click
        If d = False Then
            If txtTaijuu.Text = "" Then
                txtTaijuu.Text = "0.00"
                txtTaijuu.SelectionStart = 1
            Else
                txtTaijuu.SelectionStart = txtTaijuu.Text.Length - 3
            End If
        ElseIf d = True Then
            Dim tb As TextBox = CType(sender, TextBox)
            Dim currentSelectionStart As Integer = tb.SelectionStart
            Dim maxSelectionStart As Integer = tb.Text.Length
            If currentSelectionStart = maxSelectionStart Then
                tb.Select(currentSelectionStart - 1, 1)
            ElseIf currentSelectionStart = maxSelectionStart - 2 OrElse currentSelectionStart = maxSelectionStart - 1 Then
                tb.Select(currentSelectionStart, 1)
            End If
        End If

        d = True
    End Sub

    Private Sub txtTaijuu_Enter(sender As Object, e As System.EventArgs) Handles txtTaijuu.Enter
        d = False
    End Sub

    Private Sub TaijuuTextBox_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtTaijuu.KeyDown
        Dim tb As TextBox = CType(sender, TextBox)
        Dim inputIntStr As String = tb.Text.Split(".")(0) '整数部
        Dim inputDecimalStr As String = tb.Text.Split(".")(1) '小数部
        Dim maxSelectionStart As Integer = tb.Text.Length '最大値
        Dim currentSelectionStart As Integer = tb.SelectionStart '現在選択位置
        Dim decimalPointSelectionStart As Integer = maxSelectionStart - 3 ' 小数点の位置

        If e.KeyCode = Windows.Forms.Keys.Left Then
            'カーソルを左に動かす
            If currentSelectionStart <> 0 Then
                tb.SelectionStart -= 1
                If currentSelectionStart = maxSelectionStart - 1 OrElse currentSelectionStart = maxSelectionStart Then
                    tb.SelectionStart = maxSelectionStart - 2
                    tb.Select(tb.SelectionStart, 1)
                Else
                    tb.SelectionLength = 0
                End If
            End If
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Windows.Forms.Keys.Right Then
            'カーソルを右に動かす
            If currentSelectionStart <> maxSelectionStart Then
                tb.SelectionStart += 1
                If currentSelectionStart = decimalPointSelectionStart Then
                    tb.Select(tb.SelectionStart, 1)
                ElseIf currentSelectionStart = maxSelectionStart - 2 OrElse currentSelectionStart = maxSelectionStart - 1 Then
                    tb.Select(maxSelectionStart - 1, 1)
                End If
            End If
            e.SuppressKeyPress = True
        ElseIf (Keys.NumPad0 <= e.KeyCode AndAlso e.KeyCode <= Keys.NumPad9) OrElse (Keys.D0 <= e.KeyCode AndAlso e.KeyCode <= Keys.D9) Then
            Dim keyDownChar As String = If(Keys.NumPad0 <= e.KeyCode, Chr(e.KeyCode - 48), Chr(e.KeyCode))
            '数字の入力
            If currentSelectionStart > decimalPointSelectionStart Then
                '小数部の入力
                If currentSelectionStart = maxSelectionStart - 2 Then
                    '小数第一位
                    tb.Text = inputIntStr & "." & keyDownChar & inputDecimalStr.Substring(1, 1)
                    tb.SelectionStart = currentSelectionStart + 1
                    tb.Select(tb.SelectionStart, 1)
                ElseIf currentSelectionStart = maxSelectionStart - 1 OrElse currentSelectionStart = maxSelectionStart Then
                    '小数第二位
                    tb.Text = inputIntStr & "." & inputDecimalStr.Substring(0, 1) & keyDownChar
                    tb.SelectionStart = currentSelectionStart
                    tb.Select(tb.SelectionStart, 1)
                End If
            Else
                '整数部の入力
                If inputIntStr.Length < 3 Then
                    If inputIntStr = "0" Then
                        tb.Text = keyDownChar & "." & inputDecimalStr
                        tb.SelectionStart = 1
                    Else
                        If Not (currentSelectionStart = 0 AndAlso keyDownChar = "0") Then
                            tb.Text = inputIntStr.Insert(currentSelectionStart, keyDownChar) & "." & inputDecimalStr
                            tb.SelectionStart = currentSelectionStart + 1
                        End If
                    End If
                End If
            End If
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Decimal OrElse e.KeyCode = Keys.OemPeriod Then
            '小数点の入力
            If currentSelectionStart = decimalPointSelectionStart Then
                tb.SelectionStart += 1
                tb.Select(tb.SelectionStart, 1)
            End If
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Back Then
            Dim selectionLength As Integer = tb.SelectionLength
            Dim selectionEnd As Integer = currentSelectionStart + selectionLength

            If selectionLength = tb.Text.Length Then
                '全選択の場合
                tb.Text = "0.00"
                tb.SelectionStart = 1
            ElseIf selectionEnd > decimalPointSelectionStart + 1 Then
                '小数部分も選択している状態
                If selectionLength = 1 Then
                    '選択部分が一か所の場合
                    If currentSelectionStart = maxSelectionStart - 2 Then
                        '小数第一位
                        tb.Text = inputIntStr & "." & "0" & inputDecimalStr.Substring(1, 1)
                        tb.SelectionStart = currentSelectionStart - 1
                        tb.SelectionLength = 0
                    ElseIf currentSelectionStart = maxSelectionStart - 1 OrElse currentSelectionStart = maxSelectionStart Then
                        '小数第二位
                        tb.Text = inputIntStr & "." & inputDecimalStr.Substring(0, 1) & "0"
                        tb.SelectionStart = decimalPointSelectionStart + 1
                        tb.Select(tb.SelectionStart, 1)
                    End If
                Else
                    If selectionEnd = maxSelectionStart - 1 Then
                        '小数第一位
                        tb.Text = inputIntStr & "." & "0" & inputDecimalStr.Substring(1, 1)
                        tb.Select(selectionEnd - 1, 1)
                    Else
                        '小数第二位
                        tb.Text = inputIntStr & "." & inputDecimalStr.Substring(0, 1) & "0"
                        tb.Select(selectionEnd - 1, 1)
                    End If
                End If
            Else
                '小数部分が選ばれていない場合
                If selectionLength = 0 Then
                    '選択部分がない場合
                    If currentSelectionStart <> 0 Then
                        If inputIntStr.Length = 1 AndAlso currentSelectionStart = decimalPointSelectionStart Then
                            tb.Text = "0." & inputDecimalStr
                            tb.SelectionStart = currentSelectionStart
                        Else
                            tb.Text = inputIntStr.Remove(currentSelectionStart - 1, 1) & "." & inputDecimalStr
                            tb.SelectionStart = currentSelectionStart - 1
                        End If
                    End If
                Else
                    '選択されている場合

                    '小数点も選択されている場合
                    If selectionEnd > decimalPointSelectionStart Then
                        selectionLength -= 1
                    End If

                    If selectionLength = inputIntStr.Length Then
                        '整数部が全選択の場合
                        tb.Text = "0." & inputDecimalStr
                        tb.SelectionStart = 1
                    Else
                        tb.Text = inputIntStr.Remove(currentSelectionStart, selectionLength) & "." & inputDecimalStr
                        tb.SelectionStart = currentSelectionStart
                    End If
                End If
            End If
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Delete Then    'Delキーが押されたとき
            Dim selectionLength As Integer = tb.SelectionLength
            Dim selectionEnd As Integer = currentSelectionStart + selectionLength

            If Val(txtTaijuu.Text) = 0 Then     'テキストの中身が0のとき
                e.SuppressKeyPress = True   '何もしない
            ElseIf Val(txtTaijuu.Text) <> 0 AndAlso txtTaijuu.SelectionLength = maxSelectionStart Then     'テキストの中身が0ではない且つ文字を全選択しているとき
                txtTaijuu.Text = "0.00"     '値を0.00にする
                e.SuppressKeyPress = True
            ElseIf selectionEnd > decimalPointSelectionStart + 1 Then
                '小数部分も選択している状態
                If selectionLength = 1 Then
                    '選択部分が一か所の場合
                    If currentSelectionStart = maxSelectionStart - 2 Then
                        '小数第一位
                        tb.Text = inputIntStr & "." & "0" & Strings.Right(inputDecimalStr, 1)
                        tb.SelectionStart = currentSelectionStart
                        tb.Select(tb.SelectionStart, 0)
                        tb.SelectionLength = 1
                        e.SuppressKeyPress = True
                    ElseIf currentSelectionStart = maxSelectionStart - 1 OrElse currentSelectionStart = maxSelectionStart Then
                        '小数第二位
                        tb.Text = inputIntStr & "." & Strings.Left(inputDecimalStr, 1) & "0"
                        tb.SelectionStart = decimalPointSelectionStart + 2
                        tb.Select(tb.SelectionStart, 1)
                        e.SuppressKeyPress = True
                    End If
                Else
                    If selectionEnd = maxSelectionStart - 1 Then
                        '小数第一位
                        tb.Text = inputIntStr & "." & "0" & inputDecimalStr.Substring(1, 1)
                        tb.Select(selectionEnd - 1, 1)
                        e.SuppressKeyPress = True
                    Else
                        '小数第二位
                        tb.Text = inputIntStr & "." & inputDecimalStr.Substring(0, 1) & "0"
                        tb.Select(selectionEnd - 1, 1)
                        e.SuppressKeyPress = True
                    End If
                End If
            Else
                '小数部分が選ばれていない場合
                If selectionLength = 0 Then
                    '選択部分がない場合
                    If currentSelectionStart <> decimalPointSelectionStart Then     '小数点の前の位置以外でDelを押されたとき
                        If inputIntStr.Length = 1 Then
                            tb.Text = "0." & inputDecimalStr
                            tb.SelectionStart = currentSelectionStart
                            e.SuppressKeyPress = True
                        Else
                            'Delの処理
                        End If
                    Else    '小数点の前の位置でDelを押されたとき
                        e.SuppressKeyPress = True   'なにもしない
                    End If
                Else
                    '選択されている場合
                    '小数点も選択されている場合
                    If selectionEnd > decimalPointSelectionStart Then
                        selectionLength -= 1
                    End If
                    If selectionLength = inputIntStr.Length Then
                        '整数部が全選択の場合
                        tb.Text = "0." & inputDecimalStr
                        tb.SelectionStart = 1
                    Else
                        tb.Text = inputIntStr.Remove(currentSelectionStart, selectionLength) & "." & inputDecimalStr
                        tb.SelectionStart = currentSelectionStart
                    End If
                    e.SuppressKeyPress = True
                End If
            End If
        Else
            e.SuppressKeyPress = True
        End If
    End Sub

    
End Class