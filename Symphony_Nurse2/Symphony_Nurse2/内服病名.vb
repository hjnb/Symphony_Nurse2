Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Public Class 内服病名

    Private Sub 内服病名_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim forward As Boolean = e.Modifiers <> Keys.Shift
            Me.SelectNextControl(Me.ActiveControl, forward, True, True, True)
            e.Handled = True
        End If
    End Sub

    Private Sub 内服病名_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
        Me.DesktopLocation = New Point(0, 50)

        With DataGridView1
            .RowTemplate.Height = 18
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .RowHeadersVisible = False '行ヘッダー削除
            .ColumnHeadersVisible = False '列ヘッダー削除
            .ReadOnly = True '編集禁止
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        End With

        With DataGridView2
            .RowTemplate.Height = 16
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            '.RowHeadersVisible = False '行ヘッダー削除
            .ColumnHeadersVisible = False '列ヘッダー削除
            .ReadOnly = True '編集禁止
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        End With

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        SQLCm.CommandText = "select * from Unit order by Gyo"
        Adapter.Fill(Table)
        DataGridView3.DataSource = Table

        With DataGridView1
            .Columns.Add("RoomNo", "　")
            .Columns.Add("Nam", "氏名")
            .Columns(0).Width = 30
            .Columns(1).Width = 100
        End With

        '名前をセット
        For n As Integer = 1 To 10
            For i As Integer = 0 To 9
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i + 10 * (n - 1)).Cells(1).Value = DataGridView3(3 * n - 1, i).Value
            Next
        Next

        '部屋番号をセット
        Dim Ha As Integer = 1
        Dim Hb As Integer = 1
        Dim Hc As Integer = 1
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            For C As Integer = Hc To 2       '本館とアネックス分
                For A As Integer = Ha To 6       '部屋番号の頭
                    If A = 4 Then
                        '何もしない
                        Ha = Ha + 1
                    ElseIf A = 6 Then
                        For B As Integer = Hb To 12      '部屋番号の後ろ
                            If B = 4 Then
                                '何もしない
                                Hb = Hb + 1
                                'GoTo line1
                            ElseIf B = 9 Then
                                '何もしない
                                Hb = Hb + 1
                                'GoTo line1
                            ElseIf B < 10 Then      'Bが一桁の時
                                DataGridView1.Rows(i).Cells(0).Value = A & "0" & B
                                Hb = Hb + 1
                                GoTo line1
                            ElseIf B = 12 Then
                                DataGridView1.Rows(i).Cells(0).Value = A & B
                                'Ha = Ha + 1
                                Hb = 1
                                Ha = 1
                                Hc = Hc + 1
                                GoTo line1
                            ElseIf B > 9 Then       'Bが二桁の時
                                DataGridView1.Rows(i).Cells(0).Value = A & B
                                Hb = Hb + 1
                                GoTo line1
                            End If
                        Next
                    Else
                        For B As Integer = Hb To 12      '部屋番号の後ろ
                            If B = 4 Then
                                '何もしない
                                Hb = Hb + 1
                                'GoTo line1
                            ElseIf B = 9 Then
                                '何もしない
                                Hb = Hb + 1
                                'GoTo line1
                            ElseIf B < 10 Then      'Bが一桁の時
                                DataGridView1.Rows(i).Cells(0).Value = A & "0" & B
                                Hb = Hb + 1
                                GoTo line1
                            ElseIf B = 12 Then
                                DataGridView1.Rows(i).Cells(0).Value = A & B
                                Ha = Ha + 1
                                Hb = 1
                                GoTo line1
                            ElseIf B > 9 Then       'Bが二桁の時
                                DataGridView1.Rows(i).Cells(0).Value = A & B
                                Hb = Hb + 1
                                GoTo line1
                            End If
                        Next
                    End If
                Next
            Next
line1:
        Next

        'ユニット名の追加と設定
        With DataGridView1
            .Rows.Insert(0)
            .Rows(0).Cells(1).Value = "空"
            .Rows(0).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(11)
            .Rows(11).Cells(1).Value = "森"
            .Rows(11).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(22)
            .Rows(22).Cells(1).Value = "星"
            .Rows(22).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(33)
            .Rows(33).Cells(1).Value = "月"
            .Rows(33).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(44)
            .Rows(44).Cells(1).Value = "花"
            .Rows(44).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(55)
            .Rows(55).Cells(1).Value = "丘"
            .Rows(55).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(66)
            .Rows(66).Cells(1).Value = "虹"
            .Rows(66).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(77)
            .Rows(77).Cells(1).Value = "光"
            .Rows(77).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(88)
            .Rows(88).Cells(1).Value = "雪"
            .Rows(88).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(99)
            .Rows(99).Cells(1).Value = "風"
            .Rows(99).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
        End With

        lblID.Text = ""

        KeyPreview = True

        rbnSubete.Checked = True

        Dim SQLCm5 As OleDbCommand = Cn.CreateCommand
        Dim Adapter5 As New OleDbDataAdapter(SQLCm5)
        Dim Table5 As New DataTable

        SQLCm5.CommandText = "select Md1, Md2, Md4 from BshtMd order by Md1, Md2, Md4"
        Adapter5.Fill(Table5)
        DataGridView5.DataSource = Table5

        lblHeya.Text = ""
        lblName.Text = ""
        YmdBox1.setADStr(Today.ToString("yyyy/MM/dd"))
    End Sub

    Private Function NullCheck(cellvalue As Object) As String
        Return If(IsDBNull(cellvalue), "", cellvalue)
    End Function

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        btnKuria.PerformClick()
        DataGridView4.Columns.Clear()
        Dim s As Integer = DataGridView1.CurrentRow.Index
        lblName.Text = NullCheck(DataGridView1(1, s).Value)
        Dim DGV3rowcount As Integer = DataGridView3.Rows.Count
        If DataGridView1(0, s).Value <> "" Then         '部屋番号がある行を選択したとき
            If s < 11 Then
                lblHeya.Text = "空　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(2, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(1, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 22 Then
                lblHeya.Text = "森　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(5, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(4, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 33 Then
                lblHeya.Text = "星　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(8, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(7, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 44 Then
                lblHeya.Text = "月　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(11, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(10, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 55 Then
                lblHeya.Text = "花　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(14, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(13, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 66 Then
                lblHeya.Text = "丘　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(17, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(16, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 77 Then
                lblHeya.Text = "虹　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(20, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(19, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 88 Then
                lblHeya.Text = "光　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(23, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(22, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 99 Then
                lblHeya.Text = "雪　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(26, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(25, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 110 Then
                lblHeya.Text = "風　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(29, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(28, i).Value)
                        Exit For
                    End If
                Next
            End If
        Else         '部屋番号がない行を選択したとき
            lblHeya.Text = ""
            lblID.Text = ""
            lblName.Text = ""
            '何もしない
        End If

        If lblID.Text <> "" Then
            Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
            Dim SQLCm As OleDbCommand = Cn.CreateCommand
            Dim Adapter As New OleDbDataAdapter(SQLCm)
            Dim Table As New DataTable
            SQLCm.CommandText = "select * from BSht WHERE Id = " & lblID.Text & " AND Nam = '" & lblName.Text & "' order by Gyo"
            Adapter.Fill(Table)
            DataGridView4.DataSource = Table

            'データがあったら表示
            If DataGridView4.Rows.Count > 15 Then
                For Gyo As Integer = 1 To 16
                    If Gyo < 5 Then
                        If IsDBNull(NullCheck(DataGridView4(5, Gyo - 1).Value)) = False Then
                            Controls("txtByoumei" & Gyo).Text = NullCheck(DataGridView4(5, Gyo - 1).Value)
                        End If
                        If IsDBNull(NullCheck(DataGridView4(10, Gyo - 1).Value)) = False Then
                            Controls("txtTokki" & Gyo).Text = NullCheck(DataGridView4(10, Gyo - 1).Value)
                        End If
                    End If
                    Controls("txtNaihuku" & Gyo).Text = NullCheck(DataGridView4(6, Gyo - 1).Value)
                    Controls("txtRyou" & Gyo).Text = NullCheck(DataGridView4(7, Gyo - 1).Value)
                    Controls("txtKatati" & Gyo).Text = NullCheck(DataGridView4(8, Gyo - 1).Value)
                    Controls("txtJikann" & Gyo).Text = NullCheck(DataGridView4(9, Gyo - 1).Value)
                Next
                YmdBox1.setADStr(NullCheck(DataGridView4(11, 0).Value))
            End If
        End If

    End Sub

    Private Sub DataGridView2_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.CellMouseClick
        Xclear()
        Dim x As Integer = DataGridView2.CurrentRow.Index
        If rbnKigou.Checked = True Then
            txtKatatiX.Text = NullCheck(DataGridView2(1, x).Value)
        ElseIf rbnKigou.Checked = False Then
            txtNaihukuX.Text = NullCheck(DataGridView2(0, x).Value)
            txtRyouX.Text = NullCheck(DataGridView2(1, x).Value)
            txtJikannX.Text = NullCheck(DataGridView2(2, x).Value)
        End If
    End Sub

    Private Sub DataGridView2_CellPainting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles DataGridView2.CellPainting
        '列ヘッダを対象外しておかないと列ヘッダにも番号を表示してしまう

        If e.ColumnIndex < 0 And e.RowIndex >= 0 Then

            'セルを描画する

            e.Paint(e.ClipBounds, DataGridViewPaintParts.All)

            '行番号を描画する範囲を決定する

            Dim idxRect As Rectangle = e.CellBounds

            'e.Padding(値を表示する境界線からの間隔)を考慮して描画位置を決める

            Dim rectHeight As Long = e.CellStyle.Padding.Top

            Dim rectLeft As Long = e.CellStyle.Padding.Left

            idxRect.Inflate(rectLeft, rectHeight)

            '行番号を描画する

            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), e.CellStyle.Font, idxRect, e.CellStyle.ForeColor, TextFormatFlags.Right Or TextFormatFlags.VerticalCenter)

            '描画完了の通知

            e.Handled = True

        End If

    End Sub

    Private Sub btnKuria_Click(sender As System.Object, e As System.EventArgs) Handles btnKuria.Click
        lblHeya.Text = ""
        lblID.Text = ""
        lblName.Text = ""
        For i As Integer = 1 To 4
            Controls("txtByoumei" & i).Text = ""
            Controls("txtTokki" & i).Text = ""
        Next
        For i As Integer = 1 To 16
            Controls("txtNaihuku" & i).Text = ""
            Controls("txtRyou" & i).Text = ""
            Controls("txtKatati" & i).Text = ""
            Controls("txtJikann" & i).Text = ""
        Next
        YmdBox1.clearText()
    End Sub

    Private Sub btnTouroku_Click(sender As System.Object, e As System.EventArgs) Handles btnTouroku.Click
        If YmdBox1.getADStr() = "" Then
            MsgBox("日付を入力してください。")
            Return
        ElseIf lblID.Text = "" Then
            MsgBox("利用者を選択してください。")
            Return
        End If

        If DataGridView4.Rows.Count <> 0 Then
            'データがあったら変更
            Hennkou()
            btnKousinn.PerformClick()
        Else
            'データがなかったら新規追加
            Tuika()
            btnKousinn.PerformClick()
        End If

    End Sub
    Private Sub Hennkou()
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim SQL As String = ""
        SQL = "DELETE FROM BSht WHERE Id = " & lblID.Text
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
        Dim Table As DataTable = DirectCast(DataGridView4.DataSource, DataTable)
        Dim Row As DataRow = Table.NewRow

        Cn.Open()
        For Gyo As Integer = 1 To 16
            If Gyo < 5 Then
                Row("Bmei") = Controls("txtByoumei" & Gyo).Text
                Row("Tokki") = Controls("txtTokki" & Gyo).Text
            Else
                Row("Bmei") = ""
                Row("Tokki") = ""
            End If
            Row("Id") = lblID.Text
            Row("Nam") = lblName.Text
            Row("Unt") = Strings.Left(lblHeya.Text, 1)
            Row("Rm") = Strings.Right(lblHeya.Text, 3)
            Row("Gyo") = Gyo
            Row("Md1") = Controls("txtNaihuku" & Gyo).Text
            Row("Md2") = Controls("txtRyou" & Gyo).Text
            Row("Md3") = Controls("txtKatati" & Gyo).Text
            Row("Md4") = Controls("txtJikann" & Gyo).Text
            Row("Ymd") = YmdBox1.getADStr()
            Dim SQL As String = ""
            SQL = "INSERT INTO BSht (Id, Nam, Unt, Rm, Gyo, Bmei, Md1, Md2, Md3, Md4, Tokki, Ymd) VALUES ("
            SQL &= Row("Id") & ", "
            SQL &= "'" & Row("Nam") & "', "
            SQL &= "'" & Row("Unt") & "', "
            SQL &= "'" & Row("Rm") & "', "
            SQL &= "'" & Row("Gyo") & "', "
            SQL &= "'" & Row("Bmei") & "', "
            SQL &= "'" & Row("Md1") & "', "
            SQL &= "'" & Row("Md2") & "', "
            SQL &= "'" & Row("Md3") & "', "
            SQL &= "'" & Row("Md4") & "', "
            SQL &= "'" & Row("Tokki") & "', "
            SQL &= "'" & Row("Ymd") & "' "
            SQL &= ")"

            SQLCm.CommandText = SQL
            SQLCm.ExecuteNonQuery()
        Next

        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()

        'オート保存の部分
        AutoSave()
    End Sub
    Private Sub AutoSave()
        '内服のtextboxを見る
        For Naihuku As Integer = 1 To 16
            If Controls("txtNaihuku" & Naihuku).Text <> "" Then     'textboxが空じゃないとき
                For i As Integer = 0 To DataGridView5.Rows.Count - 1
                    If DataGridView5(0, i).Value = Controls("txtNaihuku" & Naihuku).Text AndAlso DataGridView5(1, i).Value = Controls("txtRyou" & Naihuku).Text AndAlso DataGridView5(2, i).Value = Controls("txtJikann" & Naihuku).Text Then  '項目名が一緒
                        GoTo line1  '同じものが見つかったら次のtextboxを見る
                    End If
                Next
                NaihukuTuika(Naihuku)
            Else        'textboxが空のとき
                'なにもしない
            End If
line1:
        Next
    End Sub
    Private Sub NaihukuTuika(NaihukuNo As Integer)
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Table As DataTable = DirectCast(DataGridView5.DataSource, DataTable)
        Dim Row As DataRow = Table.NewRow

        Cn.Open()

        Row("Md1") = Controls("txtNaihuku" & NaihukuNo).Text
        Row("Md2") = Controls("txtRyou" & NaihukuNo).Text
        Row("Md4") = Controls("txtJikann" & NaihukuNo).Text
        Dim SQL As String = ""
        SQL = "INSERT INTO BShtMd (Md1, Md2, Md4) VALUES ("
        SQL &= "'" & Row("Md1") & "', "
        SQL &= "'" & Row("Md2") & "', "
        SQL &= "'" & Row("Md4") & "' "
        SQL &= ")"
        SQLCm.CommandText = SQL
        SQLCm.ExecuteNonQuery()

        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()

    End Sub

    Private Function getNowWarekiTime() As String
        Dim NEXT_ERA_CHAR As String = "X"
        Dim eraText As String = ""
        Dim monthText As String = ""
        Dim dateText As String = ""
        Dim now As DateTime = DateTime.Now
        Dim year As Integer = now.Year
        Dim month As Integer = now.Month
        Dim day As Integer = now.Day
        'Dim time As String = now.ToString("HH:mm")

        If year = 2018 Then
            eraText = "H30"
        ElseIf year = 2019 Then
            If month >= 5 Then
                eraText = NEXT_ERA_CHAR & "01"
            Else
                eraText = "H31"
            End If
        ElseIf year >= 2020 Then
            eraText = NEXT_ERA_CHAR & If(year - 2018 < 10, "0" & (year - 2018), year - 2018)
        End If

        monthText = If(month < 10, "0" & month, month)
        dateText = If(day < 10, "0" & day, day)

        Return eraText & "/" & monthText & "/" & dateText
    End Function

    Private Sub btnInnsatu_Click(sender As System.Object, e As System.EventArgs) Handles btnInnsatu.Click
        ProgressBar1.Visible = True
        ProgressBar1.Maximum = 10
        ProgressBar1.Minimum = 0
        Label2.Visible = True

        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheets As Object
        Dim oSheet As Object
        Dim day As DateTime = DateTime.Today

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open("\\PRIMERGYTX100S1\Hakojun\事務\さかもと\Symphony_Nurse2\Nurse2.xls")
        oSheets = objWorkBook.Worksheets
        oSheet = objWorkBook.Worksheets("内服病名新")

        For i As Integer = 85 To 850 Step 85
            oSheet.Range("N" & i).Value = getNowWarekiTime()
        Next

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim id As Integer
        Dim DGV3rowcount As Integer = DataGridView3.Rows.Count
        Dim c As Integer = 8

        '利用者マスタのデータを表示
        Dim SQLCm7 As OleDbCommand = Cn.CreateCommand
        Dim Adapter7 As New OleDbDataAdapter(SQLCm7)
        Dim Table7 As New DataTable
        SQLCm7.CommandText = "select Id, Nam, Kana, Sex, Format(Birth, 'gee/mm/dd') as Birth, Int((Format(NOW(),'YYYYMMDD')-Format(Birth, 'YYYYMMDD'))/10000) as Age, Kaigo, Dsp from KihonM order by Kana"
        Adapter7.Fill(Table7)
        DataGridView7.DataSource = Table7

        For count As Integer = 0 To DGV3rowcount - 1    'ユニット居室のデータベース
            For name As Integer = c To 80 Step 8
                '空の家
                oSheet.Range("C" & name).Value = DataGridView3(2, count).Value

                If DataGridView3(2, count).Value = "" Then
                    id = 0
                Else
                    id = DataGridView3(1, count).Value
                End If

                Dim TableSora As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from BSht WHERE Id = " & id & "AND Nam = '" & DataGridView3(2, count).Value & "' order by Gyo"
                Adapter.Fill(TableSora)
                DataGridView6.DataSource = TableSora

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If DataGridView6(4, r).Value = 1 Then
                        oSheet.Range("G" & name - 3).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 2 Then
                        oSheet.Range("G" & name - 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 3 Then
                        oSheet.Range("G" & name + 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 4 Then
                        For i As Integer = 0 To DataGridView7.Rows.Count - 1
                            If DataGridView3(2, count).Value = DataGridView7(1, i).Value Then
                                oSheet.Range("G" & name + 3).Value = DataGridView7(4, i).Value & " (" & DataGridView7(5, i).Value & "歳) 要介護" & DataGridView7(6, i).Value
                                Exit For
                            End If
                        Next
                        oSheet.Range("I" & name).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 5 Then
                        oSheet.Range("I" & name + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 6 Then
                        oSheet.Range("I" & name + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 7 Then
                        oSheet.Range("I" & name + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 8 Then
                        oSheet.Range("I" & name + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + 4).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 9 Then
                        oSheet.Range("N" & name - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 10 Then
                        oSheet.Range("N" & name - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 11 Then
                        oSheet.Range("N" & name - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 12 Then
                        oSheet.Range("N" & name).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 13 Then
                        oSheet.Range("N" & name + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 14 Then
                        oSheet.Range("N" & name + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 15 Then
                        oSheet.Range("N" & name + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 16 Then
                        oSheet.Range("N" & name + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + 4).Value = DataGridView6(9, r).Value
                    End If
                Next

                '森の家
                oSheet.Range("C" & name + (85 * 1)).Value = DataGridView3(5, count).Value

                If DataGridView3(5, count).Value = "" Then
                    id = 0
                Else
                    id = DataGridView3(4, count).Value
                End If

                Dim TableMori As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from BSht WHERE Id = " & id & "AND Nam = '" & DataGridView3(5, count).Value & "' order by Gyo"
                Adapter.Fill(TableMori)
                DataGridView6.DataSource = TableMori

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If DataGridView6(4, r).Value = 1 Then
                        oSheet.Range("G" & name + (85 * 1) - 3).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 1) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 1) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 1) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 1) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 2 Then
                        oSheet.Range("G" & name + (85 * 1) - 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 1) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 1) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 1) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 1) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 3 Then
                        oSheet.Range("G" & name + (85 * 1) + 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 1) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 1) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 1) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 1) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 4 Then
                        For i As Integer = 0 To DataGridView7.Rows.Count - 1
                            If DataGridView3(5, count).Value = DataGridView7(1, i).Value Then
                                oSheet.Range("G" & name + (85 * 1) + 3).Value = DataGridView7(4, i).Value & " (" & DataGridView7(5, i).Value & "歳) 要介護" & DataGridView7(6, i).Value
                                Exit For
                            Else
                                oSheet.Range("G" & name + (85 * 1) + 3).Value = DataGridView6(5, r).Value
                            End If
                        Next
                        oSheet.Range("I" & name + (85 * 1)).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 1)).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 1)).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 1)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 5 Then
                        oSheet.Range("I" & name + (85 * 1) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 1) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 1) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 1) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 6 Then
                        oSheet.Range("I" & name + (85 * 1) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 1) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 1) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 1) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 7 Then
                        oSheet.Range("I" & name + (85 * 1) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 1) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 1) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 1) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 8 Then
                        oSheet.Range("I" & name + (85 * 1) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 1) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 1) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 1) + 4).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 9 Then
                        oSheet.Range("N" & name + (85 * 1) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 1) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 1) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 1) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 10 Then
                        oSheet.Range("N" & name + (85 * 1) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 1) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 1) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 1) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 11 Then
                        oSheet.Range("N" & name + (85 * 1) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 1) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 1) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 1) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 12 Then
                        oSheet.Range("N" & name + (85 * 1)).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 1)).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 1)).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 1)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 13 Then
                        oSheet.Range("N" & name + (85 * 1) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 1) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 1) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 1) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 14 Then
                        oSheet.Range("N" & name + (85 * 1) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 1) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 1) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 1) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 15 Then
                        oSheet.Range("N" & name + (85 * 1) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 1) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 1) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 1) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 16 Then
                        oSheet.Range("N" & name + (85 * 1) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 1) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 1) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 1) + 4).Value = DataGridView6(9, r).Value
                    End If
                Next

                '星の家
                oSheet.Range("C" & name + (85 * 2)).Value = DataGridView3(8, count).Value

                If DataGridView3(8, count).Value = "" Then
                    id = 0
                Else
                    id = DataGridView3(7, count).Value
                End If

                Dim TableHosi As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from BSht WHERE Id = " & id & "AND Nam = '" & DataGridView3(8, count).Value & "' order by Gyo"
                Adapter.Fill(TableHosi)
                DataGridView6.DataSource = TableHosi

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If DataGridView6(4, r).Value = 1 Then
                        oSheet.Range("G" & name + (85 * 2) - 3).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 2) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 2) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 2) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 2) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 2 Then
                        oSheet.Range("G" & name + (85 * 2) - 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 2) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 2) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 2) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 2) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 3 Then
                        oSheet.Range("G" & name + (85 * 2) + 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 2) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 2) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 2) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 2) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 4 Then
                        For i As Integer = 0 To DataGridView7.Rows.Count - 1
                            If DataGridView3(8, count).Value = DataGridView7(1, i).Value Then
                                oSheet.Range("G" & name + (85 * 2) + 3).Value = DataGridView7(4, i).Value & " (" & DataGridView7(5, i).Value & "歳) 要介護" & DataGridView7(6, i).Value
                                Exit For
                            Else
                                oSheet.Range("G" & name + (85 * 2) + 3).Value = DataGridView6(5, r).Value
                            End If
                        Next
                        oSheet.Range("I" & name + (85 * 2)).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 2)).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 2)).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 2)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 5 Then
                        oSheet.Range("I" & name + (85 * 2) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 2) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 2) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 2) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 6 Then
                        oSheet.Range("I" & name + (85 * 2) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 2) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 2) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 2) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 7 Then
                        oSheet.Range("I" & name + (85 * 2) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 2) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 2) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 2) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 8 Then
                        oSheet.Range("I" & name + (85 * 2) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 2) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 2) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 2) + 4).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 9 Then
                        oSheet.Range("N" & name + (85 * 2) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 2) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 2) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 2) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 10 Then
                        oSheet.Range("N" & name + (85 * 2) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 2) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 2) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 2) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 11 Then
                        oSheet.Range("N" & name + (85 * 2) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 2) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 2) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 2) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 12 Then
                        oSheet.Range("N" & name + (85 * 2)).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 2)).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 2)).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 2)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 13 Then
                        oSheet.Range("N" & name + (85 * 2) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 2) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 2) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 2) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 14 Then
                        oSheet.Range("N" & name + (85 * 2) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 2) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 2) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 2) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 15 Then
                        oSheet.Range("N" & name + (85 * 2) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 2) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 2) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 2) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 16 Then
                        oSheet.Range("N" & name + (85 * 2) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 2) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 2) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 2) + 4).Value = DataGridView6(9, r).Value
                    End If
                Next

                '月の家
                oSheet.Range("C" & name + (85 * 3)).Value = DataGridView3(11, count).Value

                If DataGridView3(11, count).Value = "" Then
                    id = 0
                Else
                    id = DataGridView3(10, count).Value
                End If

                Dim Tabletuki As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from BSht WHERE Id = " & id & "AND Nam = '" & DataGridView3(11, count).Value & "' order by Gyo"
                Adapter.Fill(Tabletuki)
                DataGridView6.DataSource = Tabletuki

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If DataGridView6(4, r).Value = 1 Then
                        oSheet.Range("G" & name + (85 * 3) - 3).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 3) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 3) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 3) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 3) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 2 Then
                        oSheet.Range("G" & name + (85 * 3) - 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 3) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 3) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 3) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 3) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 3 Then
                        oSheet.Range("G" & name + (85 * 3) + 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 3) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 3) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 3) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 3) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 4 Then
                        For i As Integer = 0 To DataGridView7.Rows.Count - 1
                            If DataGridView3(11, count).Value = DataGridView7(1, i).Value Then
                                oSheet.Range("G" & name + (85 * 3) + 3).Value = DataGridView7(4, i).Value & " (" & DataGridView7(5, i).Value & "歳) 要介護" & DataGridView7(6, i).Value
                                Exit For
                            Else
                                oSheet.Range("G" & name + (85 * 3) + 3).Value = DataGridView6(5, r).Value
                            End If
                        Next
                        oSheet.Range("I" & name + (85 * 3)).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 3)).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 3)).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 3)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 5 Then
                        oSheet.Range("I" & name + (85 * 3) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 3) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 3) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 3) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 6 Then
                        oSheet.Range("I" & name + (85 * 3) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 3) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 3) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 3) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 7 Then
                        oSheet.Range("I" & name + (85 * 3) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 3) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 3) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 3) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 8 Then
                        oSheet.Range("I" & name + (85 * 3) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 3) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 3) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 3) + 4).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 9 Then
                        oSheet.Range("N" & name + (85 * 3) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 3) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 3) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 3) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 10 Then
                        oSheet.Range("N" & name + (85 * 3) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 3) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 3) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 3) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 11 Then
                        oSheet.Range("N" & name + (85 * 3) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 3) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 3) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 3) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 12 Then
                        oSheet.Range("N" & name + (85 * 3)).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 3)).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 3)).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 3)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 13 Then
                        oSheet.Range("N" & name + (85 * 3) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 3) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 3) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 3) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 14 Then
                        oSheet.Range("N" & name + (85 * 3) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 3) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 3) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 3) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 15 Then
                        oSheet.Range("N" & name + (85 * 3) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 3) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 3) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 3) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 16 Then
                        oSheet.Range("N" & name + (85 * 3) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 3) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 3) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 3) + 4).Value = DataGridView6(9, r).Value
                    End If
                Next

                '花の家
                oSheet.Range("C" & name + (85 * 4)).Value = DataGridView3(14, count).Value

                If DataGridView3(14, count).Value = "" Then
                    id = 0
                Else
                    id = DataGridView3(13, count).Value
                End If

                Dim TableHana As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from BSht WHERE Id = " & id & "AND Nam = '" & DataGridView3(14, count).Value & "' order by Gyo"
                Adapter.Fill(TableHana)
                DataGridView6.DataSource = TableHana

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If DataGridView6(4, r).Value = 1 Then
                        oSheet.Range("G" & name + (85 * 4) - 3).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 4) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 4) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 4) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 4) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 2 Then
                        oSheet.Range("G" & name + (85 * 4) - 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 4) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 4) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 4) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 4) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 3 Then
                        oSheet.Range("G" & name + (85 * 4) + 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 4) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 4) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 4) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 4) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 4 Then
                        For i As Integer = 0 To DataGridView7.Rows.Count - 1
                            If DataGridView3(14, count).Value = DataGridView7(1, i).Value Then
                                oSheet.Range("G" & name + (85 * 4) + 3).Value = DataGridView7(4, i).Value & " (" & DataGridView7(5, i).Value & "歳) 要介護" & DataGridView7(6, i).Value
                                Exit For
                            Else
                                oSheet.Range("G" & name + (85 * 4) + 3).Value = DataGridView6(5, r).Value
                            End If
                        Next
                        oSheet.Range("I" & name + (85 * 4)).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 4)).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 4)).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 4)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 5 Then
                        oSheet.Range("I" & name + (85 * 4) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 4) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 4) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 4) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 6 Then
                        oSheet.Range("I" & name + (85 * 4) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 4) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 4) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 4) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 7 Then
                        oSheet.Range("I" & name + (85 * 4) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 4) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 4) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 4) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 8 Then
                        oSheet.Range("I" & name + (85 * 4) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 4) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 4) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 4) + 4).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 9 Then
                        oSheet.Range("N" & name + (85 * 4) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 4) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 4) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 4) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 10 Then
                        oSheet.Range("N" & name + (85 * 4) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 4) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 4) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 4) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 11 Then
                        oSheet.Range("N" & name + (85 * 4) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 4) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 4) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 4) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 12 Then
                        oSheet.Range("N" & name + (85 * 4)).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 4)).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 4)).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 4)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 13 Then
                        oSheet.Range("N" & name + (85 * 4) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 4) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 4) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 4) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 14 Then
                        oSheet.Range("N" & name + (85 * 4) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 4) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 4) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 4) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 15 Then
                        oSheet.Range("N" & name + (85 * 4) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 4) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 4) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 4) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 16 Then
                        oSheet.Range("N" & name + (85 * 4) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 4) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 4) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 4) + 4).Value = DataGridView6(9, r).Value
                    End If
                Next

                '丘の家
                oSheet.Range("C" & name + (85 * 5)).Value = DataGridView3(17, count).Value

                If DataGridView3(17, count).Value = "" Then
                    id = 0
                Else
                    id = DataGridView3(16, count).Value
                End If

                Dim TableOka As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from BSht WHERE Id = " & id & "AND Nam = '" & DataGridView3(17, count).Value & "' order by Gyo"
                Adapter.Fill(TableOka)
                DataGridView6.DataSource = TableOka

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If DataGridView6(4, r).Value = 1 Then
                        oSheet.Range("G" & name + (85 * 5) - 3).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 5) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 5) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 5) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 5) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 2 Then
                        oSheet.Range("G" & name + (85 * 5) - 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 5) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 5) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 5) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 5) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 3 Then
                        oSheet.Range("G" & name + (85 * 5) + 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 5) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 5) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 5) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 5) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 4 Then
                        For i As Integer = 0 To DataGridView7.Rows.Count - 1
                            If DataGridView3(17, count).Value = DataGridView7(1, i).Value Then
                                oSheet.Range("G" & name + (85 * 5) + 3).Value = DataGridView7(4, i).Value & " (" & DataGridView7(5, i).Value & "歳) 要介護" & DataGridView7(6, i).Value
                                Exit For
                            Else
                                oSheet.Range("G" & name + (85 * 5) + 3).Value = DataGridView6(5, r).Value
                            End If
                        Next
                        oSheet.Range("I" & name + (85 * 5)).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 5)).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 5)).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 5)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 5 Then
                        oSheet.Range("I" & name + (85 * 5) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 5) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 5) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 5) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 6 Then
                        oSheet.Range("I" & name + (85 * 5) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 5) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 5) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 5) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 7 Then
                        oSheet.Range("I" & name + (85 * 5) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 5) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 5) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 5) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 8 Then
                        oSheet.Range("I" & name + (85 * 5) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 5) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 5) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 5) + 4).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 9 Then
                        oSheet.Range("N" & name + (85 * 5) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 5) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 5) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 5) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 10 Then
                        oSheet.Range("N" & name + (85 * 5) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 5) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 5) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 5) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 11 Then
                        oSheet.Range("N" & name + (85 * 5) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 5) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 5) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 5) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 12 Then
                        oSheet.Range("N" & name + (85 * 5)).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 5)).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 5)).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 5)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 13 Then
                        oSheet.Range("N" & name + (85 * 5) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 5) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 5) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 5) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 14 Then
                        oSheet.Range("N" & name + (85 * 5) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 5) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 5) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 5) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 15 Then
                        oSheet.Range("N" & name + (85 * 5) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 5) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 5) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 5) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 16 Then
                        oSheet.Range("N" & name + (85 * 5) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 5) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 5) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 5) + 4).Value = DataGridView6(9, r).Value
                    End If
                Next

                '虹の家
                oSheet.Range("C" & name + (85 * 6)).Value = DataGridView3(20, count).Value

                If DataGridView3(20, count).Value = "" Then
                    id = 0
                Else
                    id = DataGridView3(19, count).Value
                End If

                Dim TableNiji As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from BSht WHERE Id = " & id & "AND Nam = '" & DataGridView3(20, count).Value & "' order by Gyo"
                Adapter.Fill(TableNiji)
                DataGridView6.DataSource = TableNiji

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If DataGridView6(4, r).Value = 1 Then
                        oSheet.Range("G" & name + (85 * 6) - 3).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 6) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 6) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 6) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 6) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 2 Then
                        oSheet.Range("G" & name + (85 * 6) - 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 6) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 6) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 6) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 6) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 3 Then
                        oSheet.Range("G" & name + (85 * 6) + 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 6) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 6) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 6) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 6) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 4 Then
                        For i As Integer = 0 To DataGridView7.Rows.Count - 1
                            If DataGridView3(20, count).Value = DataGridView7(1, i).Value Then
                                oSheet.Range("G" & name + (85 * 6) + 3).Value = DataGridView7(4, i).Value & " (" & DataGridView7(5, i).Value & "歳) 要介護" & DataGridView7(6, i).Value
                                Exit For
                            Else
                                oSheet.Range("G" & name + (85 * 6) + 3).Value = DataGridView6(5, r).Value
                            End If
                        Next
                        oSheet.Range("I" & name + (85 * 6)).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 6)).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 6)).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 6)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 5 Then
                        oSheet.Range("I" & name + (85 * 6) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 6) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 6) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 6) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 6 Then
                        oSheet.Range("I" & name + (85 * 6) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 6) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 6) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 6) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 7 Then
                        oSheet.Range("I" & name + (85 * 6) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 6) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 6) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 6) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 8 Then
                        oSheet.Range("I" & name + (85 * 6) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 6) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 6) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 6) + 4).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 9 Then
                        oSheet.Range("N" & name + (85 * 6) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 6) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 6) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 6) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 10 Then
                        oSheet.Range("N" & name + (85 * 6) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 6) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 6) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 6) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 11 Then
                        oSheet.Range("N" & name + (85 * 6) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 6) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 6) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 6) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 12 Then
                        oSheet.Range("N" & name + (85 * 6)).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 6)).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 6)).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 6)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 13 Then
                        oSheet.Range("N" & name + (85 * 6) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 6) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 6) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 6) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 14 Then
                        oSheet.Range("N" & name + (85 * 6) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 6) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 6) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 6) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 15 Then
                        oSheet.Range("N" & name + (85 * 6) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 6) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 6) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 6) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 16 Then
                        oSheet.Range("N" & name + (85 * 6) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 6) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 6) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 6) + 4).Value = DataGridView6(9, r).Value
                    End If
                Next

                '光の家
                oSheet.Range("C" & name + (85 * 7)).Value = DataGridView3(23, count).Value

                If DataGridView3(23, count).Value = "" Then
                    id = 0
                Else
                    id = DataGridView3(22, count).Value
                End If

                Dim TableHikari As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from BSht WHERE Id = " & id & "AND Nam = '" & DataGridView3(23, count).Value & "' order by Gyo"
                Adapter.Fill(TableHikari)
                DataGridView6.DataSource = TableHikari

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If DataGridView6(4, r).Value = 1 Then
                        oSheet.Range("G" & name + (85 * 7) - 3).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 7) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 7) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 7) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 7) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 2 Then
                        oSheet.Range("G" & name + (85 * 7) - 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 7) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 7) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 7) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 7) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 3 Then
                        oSheet.Range("G" & name + (85 * 7) + 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 7) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 7) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 7) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 7) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 4 Then
                        For i As Integer = 0 To DataGridView7.Rows.Count - 1
                            If DataGridView3(23, count).Value = DataGridView7(1, i).Value Then
                                oSheet.Range("G" & name + (85 * 7) + 3).Value = DataGridView7(4, i).Value & " (" & DataGridView7(5, i).Value & "歳) 要介護" & DataGridView7(6, i).Value
                                Exit For
                            Else
                                oSheet.Range("G" & name + (85 * 7) + 3).Value = DataGridView6(5, r).Value
                            End If
                        Next
                        oSheet.Range("I" & name + (85 * 7)).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 7)).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 7)).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 7)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 5 Then
                        oSheet.Range("I" & name + (85 * 7) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 7) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 7) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 7) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 6 Then
                        oSheet.Range("I" & name + (85 * 7) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 7) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 7) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 7) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 7 Then
                        oSheet.Range("I" & name + (85 * 7) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 7) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 7) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 7) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 8 Then
                        oSheet.Range("I" & name + (85 * 7) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 7) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 7) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 7) + 4).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 9 Then
                        oSheet.Range("N" & name + (85 * 7) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 7) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 7) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 7) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 10 Then
                        oSheet.Range("N" & name + (85 * 7) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 7) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 7) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 7) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 11 Then
                        oSheet.Range("N" & name + (85 * 7) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 7) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 7) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 7) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 12 Then
                        oSheet.Range("N" & name + (85 * 7)).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 7)).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 7)).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 7)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 13 Then
                        oSheet.Range("N" & name + (85 * 7) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 7) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 7) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 7) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 14 Then
                        oSheet.Range("N" & name + (85 * 7) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 7) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 7) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 7) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 15 Then
                        oSheet.Range("N" & name + (85 * 7) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 7) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 7) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 7) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 16 Then
                        oSheet.Range("N" & name + (85 * 7) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 7) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 7) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 7) + 4).Value = DataGridView6(9, r).Value
                    End If
                Next

                '雪の家
                oSheet.Range("C" & name + (85 * 8)).Value = DataGridView3(26, count).Value

                If DataGridView3(26, count).Value = "" Then
                    id = 0
                Else
                    id = DataGridView3(25, count).Value
                End If

                Dim TableYuki As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from BSht WHERE Id = " & id & "AND Nam = '" & DataGridView3(26, count).Value & "' order by Gyo"
                Adapter.Fill(TableYuki)
                DataGridView6.DataSource = TableYuki

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If DataGridView6(4, r).Value = 1 Then
                        oSheet.Range("G" & name + (85 * 8) - 3).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 8) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 8) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 8) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 8) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 2 Then
                        oSheet.Range("G" & name + (85 * 8) - 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 8) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 8) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 8) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 8) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 3 Then
                        oSheet.Range("G" & name + (85 * 8) + 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 8) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 8) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 8) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 8) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 4 Then
                        For i As Integer = 0 To DataGridView7.Rows.Count - 1
                            If DataGridView3(26, count).Value = DataGridView7(1, i).Value Then
                                oSheet.Range("G" & name + (85 * 8) + 3).Value = DataGridView7(4, i).Value & " (" & DataGridView7(5, i).Value & "歳) 要介護" & DataGridView7(6, i).Value
                                Exit For
                            Else
                                oSheet.Range("G" & name + (85 * 8) + 3).Value = DataGridView6(5, r).Value
                            End If
                        Next
                        oSheet.Range("I" & name + (85 * 8)).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 8)).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 8)).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 8)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 5 Then
                        oSheet.Range("I" & name + (85 * 8) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 8) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 8) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 8) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 6 Then
                        oSheet.Range("I" & name + (85 * 8) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 8) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 8) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 8) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 7 Then
                        oSheet.Range("I" & name + (85 * 8) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 8) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 8) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 8) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 8 Then
                        oSheet.Range("I" & name + (85 * 8) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 8) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 8) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 8) + 4).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 9 Then
                        oSheet.Range("N" & name + (85 * 8) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 8) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 8) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 8) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 10 Then
                        oSheet.Range("N" & name + (85 * 8) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 8) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 8) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 8) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 11 Then
                        oSheet.Range("N" & name + (85 * 8) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 8) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 8) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 8) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 12 Then
                        oSheet.Range("N" & name + (85 * 8)).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 8)).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 8)).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 8)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 13 Then
                        oSheet.Range("N" & name + (85 * 8) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 8) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 8) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 8) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 14 Then
                        oSheet.Range("N" & name + (85 * 8) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 8) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 8) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 8) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 15 Then
                        oSheet.Range("N" & name + (85 * 8) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 8) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 8) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 8) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 16 Then
                        oSheet.Range("N" & name + (85 * 8) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 8) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 8) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 8) + 4).Value = DataGridView6(9, r).Value
                    End If
                Next

                '風の家
                oSheet.Range("C" & name + (85 * 9)).Value = DataGridView3(29, count).Value

                If DataGridView3(29, count).Value = "" Then
                    id = 0
                Else
                    id = DataGridView3(28, count).Value
                End If

                Dim TableKaze As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from BSht WHERE Id = " & id & "AND Nam = '" & DataGridView3(29, count).Value & "' order by Gyo"
                Adapter.Fill(TableKaze)
                DataGridView6.DataSource = TableKaze

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If DataGridView6(4, r).Value = 1 Then
                        oSheet.Range("G" & name + (85 * 9) - 3).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 9) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 9) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 9) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 9) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 2 Then
                        oSheet.Range("G" & name + (85 * 9) - 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 9) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 9) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 9) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 9) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 3 Then
                        oSheet.Range("G" & name + (85 * 9) + 1).Value = DataGridView6(5, r).Value
                        oSheet.Range("I" & name + (85 * 9) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 9) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 9) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 9) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 4 Then
                        For i As Integer = 0 To DataGridView7.Rows.Count - 1
                            If DataGridView3(29, count).Value = DataGridView7(1, i).Value Then
                                oSheet.Range("G" & name + (85 * 9) + 3).Value = DataGridView7(4, i).Value & " (" & DataGridView7(5, i).Value & "歳) 要介護" & DataGridView7(6, i).Value
                                Exit For
                            Else
                                oSheet.Range("G" & name + (85 * 9) + 3).Value = DataGridView6(5, r).Value
                            End If
                        Next
                        oSheet.Range("I" & name + (85 * 9)).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 9)).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 9)).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 9)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 5 Then
                        oSheet.Range("I" & name + (85 * 9) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 9) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 9) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 9) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 6 Then
                        oSheet.Range("I" & name + (85 * 9) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 9) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 9) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 9) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 7 Then
                        oSheet.Range("I" & name + (85 * 9) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 9) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 9) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 9) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 8 Then
                        oSheet.Range("I" & name + (85 * 9) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("J" & name + (85 * 9) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("K" & name + (85 * 9) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("L" & name + (85 * 9) + 4).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 9 Then
                        oSheet.Range("N" & name + (85 * 9) - 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 9) - 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 9) - 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 9) - 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 10 Then
                        oSheet.Range("N" & name + (85 * 9) - 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 9) - 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 9) - 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 9) - 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 11 Then
                        oSheet.Range("N" & name + (85 * 9) - 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 9) - 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 9) - 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 9) - 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 12 Then
                        oSheet.Range("N" & name + (85 * 9)).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 9)).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 9)).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 9)).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 13 Then
                        oSheet.Range("N" & name + (85 * 9) + 1).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 9) + 1).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 9) + 1).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 9) + 1).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 14 Then
                        oSheet.Range("N" & name + (85 * 9) + 2).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 9) + 2).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 9) + 2).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 9) + 2).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 15 Then
                        oSheet.Range("N" & name + (85 * 9) + 3).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 9) + 3).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 9) + 3).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 9) + 3).Value = DataGridView6(9, r).Value
                    ElseIf DataGridView6(4, r).Value = 16 Then
                        oSheet.Range("N" & name + (85 * 9) + 4).Value = DataGridView6(6, r).Value
                        oSheet.Range("O" & name + (85 * 9) + 4).Value = DataGridView6(7, r).Value
                        oSheet.Range("P" & name + (85 * 9) + 4).Value = DataGridView6(8, r).Value
                        oSheet.Range("Q" & name + (85 * 9) + 4).Value = DataGridView6(9, r).Value
                    End If
                Next

                c = name + 8
                Exit For
            Next
            ProgressBar1.Value = count
        Next
        ProgressBar1.Value = 10


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

        ProgressBar1.Value = 0
        ProgressBar1.Visible = False
        Label2.Visible = False
    End Sub

    Private Sub btnKousinn_Click(sender As System.Object, e As System.EventArgs) Handles btnKousinn.Click
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        SQLCm.CommandText = "select * from BSht WHERE Id = " & lblID.Text & " order by Gyo"
        Adapter.Fill(Table)
        DataGridView4.DataSource = Table

        Dim SQLCm5 As OleDbCommand = Cn.CreateCommand
        Dim Adapter5 As New OleDbDataAdapter(SQLCm5)
        Dim Table5 As New DataTable
        SQLCm5.CommandText = "select Md1, Md2, Md4 from BshtMd order by Md1, Md2, Md4"
        Adapter5.Fill(Table5)
        DataGridView5.DataSource = Table5

        If rbnSubete.Checked = True Then
            rbnSubete.Checked = False
            rbnSubete.Checked = True
        ElseIf rbnKigou.Checked = True Then
            rbnKigou.Checked = False
            rbnKigou.Checked = True
        ElseIf rbnA.Checked = True Then
            rbnA.Checked = False
            rbnA.Checked = True
        ElseIf rbnK.Checked = True Then
            rbnK.Checked = False
            rbnK.Checked = True
        ElseIf rbnS.Checked = True Then
            rbnS.Checked = False
            rbnS.Checked = True
        ElseIf rbnT.Checked = True Then
            rbnT.Checked = False
            rbnT.Checked = True
        ElseIf rbnN.Checked = True Then
            rbnN.Checked = False
            rbnN.Checked = True
        ElseIf rbnH.Checked = True Then
            rbnH.Checked = False
            rbnH.Checked = True
        ElseIf rbnM.Checked = True Then
            rbnM.Checked = False
            rbnM.Checked = True
        ElseIf rbnY.Checked = True Then
            rbnY.Checked = False
            rbnY.Checked = True
        ElseIf rbnR.Checked = True Then
            rbnR.Checked = False
            rbnR.Checked = True
        ElseIf rbnW.Checked = True Then
            rbnW.Checked = False
            rbnW.Checked = True
        End If
    End Sub

    Private Sub rbnKigou_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnKigou.CheckedChanged
        DataGridView2.DataSource.Clear()
        Dim Table As New DataTable
        Table.Columns.Add("Md1")
        Table.Columns.Add("Md2")
        Table.Columns.Add("Md4")
        For i As Integer = 1 To 3
            Dim row As DataRow = Table.NewRow
            row("Md1") = ""
            row("Md2") = ""
            row("Md4") = ""
            Table.Rows.Add(row)
        Next

        DataGridView2.DataSource = Table

        DataGridView2(0, 0).Value = "-"
        DataGridView2(0, 1).Value = "-"
        DataGridView2(0, 2).Value = "-"
        DataGridView2(1, 0).Value = "┐"
        DataGridView2(1, 1).Value = "｜"
        DataGridView2(1, 2).Value = "┘"
    End Sub
    Private Sub rbn_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnSubete.CheckedChanged, rbnA.CheckedChanged, rbnK.CheckedChanged, rbnS.CheckedChanged, rbnT.CheckedChanged, rbnN.CheckedChanged, rbnH.CheckedChanged, rbnM.CheckedChanged, rbnY.CheckedChanged, rbnR.CheckedChanged, rbnW.CheckedChanged
        Dim rbn As RadioButton = CType(sender, RadioButton)
        Dim rbnName As String = rbn.Text

        DataGridView2.Columns.Clear()
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        If rbnName = "全て" Then
            SQLCm.CommandText = "select Md1, Md2, Md4 from BshtMd order by Md1, Md2, Md4"
        ElseIf rbnName = "ア" Then
            SQLCm.CommandText = "select Md1, Md2, Md4 from BshtMd WHERE Md1 LIKE '[ア-オ]%' order by Md1, Md2, Md4"
        ElseIf rbnName = "カ" Then
            SQLCm.CommandText = "select Md1, Md2, Md4 from BshtMd WHERE Md1 LIKE '[カ-コ]%' order by Md1, Md2, Md4"
        ElseIf rbnName = "サ" Then
            SQLCm.CommandText = "select Md1, Md2, Md4 from BshtMd WHERE Md1 LIKE '[サ-ソ]%' order by Md1, Md2, Md4"
        ElseIf rbnName = "タ" Then
            SQLCm.CommandText = "select Md1, Md2, Md4 from BshtMd WHERE Md1 LIKE '[タ-ト]%' order by Md1, Md2, Md4"
        ElseIf rbnName = "ナ" Then
            SQLCm.CommandText = "select Md1, Md2, Md4 from BshtMd WHERE Md1 LIKE '[ナ-ノ]%' order by Md1, Md2, Md4"
        ElseIf rbnName = "ハ" Then
            SQLCm.CommandText = "select Md1, Md2, Md4 from BshtMd WHERE Md1 LIKE '[ハ-ホ]%' order by Md1, Md2, Md4"
        ElseIf rbnName = "マ" Then
            SQLCm.CommandText = "select Md1, Md2, Md4 from BshtMd WHERE Md1 LIKE '[マ-モ]%' order by Md1, Md2, Md4"
        ElseIf rbnName = "ヤ" Then
            SQLCm.CommandText = "select Md1, Md2, Md4 from BshtMd WHERE Md1 LIKE '[ヤ-ヨ]%' order by Md1, Md2, Md4"
        ElseIf rbnName = "ラ" Then
            SQLCm.CommandText = "select Md1, Md2, Md4 from BshtMd WHERE Md1 LIKE '[ラ-ロ]%' order by Md1, Md2, Md4"
        ElseIf rbnName = "ワ" Then
            SQLCm.CommandText = "select Md1, Md2, Md4 from BshtMd WHERE Md1 LIKE '[ワ-ン]%' order by Md1, Md2, Md4"
        End If

        Adapter.Fill(Table)
        DataGridView2.DataSource = Table

        With DataGridView2
            .RowHeadersWidth = 30
            .Columns(0).Width = 90
            .Columns(1).Width = 45
            .Columns(2).Width = 67
        End With
    End Sub
    
    Private Sub Xclear()
        txtNaihukuX.Text = ""
        txtRyouX.Text = ""
        txtKatatiX.Text = ""
        txtJikannX.Text = ""
    End Sub
    Private Sub NaihukuGyo1_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku1.Click, txtRyou1.Click, txtKatati1.Click, txtJikann1.Click
        'Dim a As TextBox = CType(sender, TextBox)
        'MsgBox(a.Name)

        If txtNaihukuX.Text <> "" Then
            txtNaihuku1.Text = txtNaihukuX.Text
            txtRyou1.Text = txtRyouX.Text
            txtJikann1.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati1.Text = txtKatatiX.Text
            End If
        Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo2_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku2.Click, txtRyou2.Click, txtKatati2.Click, txtJikann2.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku2.Text = txtNaihukuX.Text
            txtRyou2.Text = txtRyouX.Text
            txtJikann2.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati2.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo3_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku3.Click, txtRyou3.Click, txtKatati3.Click, txtJikann3.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku3.Text = txtNaihukuX.Text
            txtRyou3.Text = txtRyouX.Text
            txtJikann3.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati3.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo4_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku4.Click, txtRyou4.Click, txtKatati4.Click, txtJikann4.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku4.Text = txtNaihukuX.Text
            txtRyou4.Text = txtRyouX.Text
            txtJikann4.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati4.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo5_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku5.Click, txtRyou5.Click, txtKatati5.Click, txtJikann5.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku5.Text = txtNaihukuX.Text
            txtRyou5.Text = txtRyouX.Text
            txtJikann5.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati5.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo6_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku6.Click, txtRyou6.Click, txtKatati6.Click, txtJikann6.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku6.Text = txtNaihukuX.Text
            txtRyou6.Text = txtRyouX.Text
            txtJikann6.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati6.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo7_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku7.Click, txtRyou7.Click, txtKatati7.Click, txtJikann7.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku7.Text = txtNaihukuX.Text
            txtRyou7.Text = txtRyouX.Text
            txtJikann7.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati7.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo8_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku8.Click, txtRyou8.Click, txtKatati8.Click, txtJikann8.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku8.Text = txtNaihukuX.Text
            txtRyou8.Text = txtRyouX.Text
            txtJikann8.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati8.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo9_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku9.Click, txtRyou9.Click, txtKatati9.Click, txtJikann9.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku9.Text = txtNaihukuX.Text
            txtRyou9.Text = txtRyouX.Text
            txtJikann9.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati9.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo10_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku10.Click, txtRyou10.Click, txtKatati10.Click, txtJikann10.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku10.Text = txtNaihukuX.Text
            txtRyou10.Text = txtRyouX.Text
            txtJikann10.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati10.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo11_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku11.Click, txtRyou11.Click, txtKatati11.Click, txtJikann11.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku11.Text = txtNaihukuX.Text
            txtRyou11.Text = txtRyouX.Text
            txtJikann11.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati11.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo12_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku12.Click, txtRyou12.Click, txtKatati12.Click, txtJikann12.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku12.Text = txtNaihukuX.Text
            txtRyou12.Text = txtRyouX.Text
            txtJikann12.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati12.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo13_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku13.Click, txtRyou13.Click, txtKatati13.Click, txtJikann13.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku13.Text = txtNaihukuX.Text
            txtRyou13.Text = txtRyouX.Text
            txtJikann13.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati13.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo14_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku14.Click, txtRyou14.Click, txtKatati14.Click, txtJikann14.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku14.Text = txtNaihukuX.Text
            txtRyou14.Text = txtRyouX.Text
            txtJikann14.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati14.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo15_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku15.Click, txtRyou15.Click, txtKatati15.Click, txtJikann15.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku15.Text = txtNaihukuX.Text
            txtRyou15.Text = txtRyouX.Text
            txtJikann15.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati15.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub
    Private Sub NaihukuGyo16_Click(sender As Object, e As System.EventArgs) Handles txtNaihuku16.Click, txtRyou16.Click, txtKatati16.Click, txtJikann16.Click
        If txtNaihukuX.Text <> "" Then
            txtNaihuku16.Text = txtNaihukuX.Text
            txtRyou16.Text = txtRyouX.Text
            txtJikann16.Text = txtJikannX.Text
            Xclear()
        Else
            If txtKatatiX.Text <> "" Then
                txtKatati16.Text = txtKatatiX.Text
            End If
            Xclear()
        End If
    End Sub

End Class