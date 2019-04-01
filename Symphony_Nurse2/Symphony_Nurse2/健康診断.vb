Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Public Class 健康診断

    Private Sub 健康診断_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim forward As Boolean = e.Modifiers <> Keys.Shift
            Me.SelectNextControl(Me.ActiveControl, forward, True, True, True)
            e.Handled = True
        End If
    End Sub

    Private Sub 健康診断_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
        Me.DesktopLocation = New Point(0, 55)
        Me.MinimizeBox = False
        Me.MaximizeBox = False

        TopForm.EnableDoubleBuffering(DataGridView1)

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

        With DataGridView5
            .RowTemplate.Height = 16
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .ColumnHeadersVisible = False '列ヘッダー削除
            .ReadOnly = True '編集禁止
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        End With

        Dim SQLCm5 As OleDbCommand = Cn.CreateCommand
        Dim Adapter5 As New OleDbDataAdapter(SQLCm5)
        Dim Table5 As New DataTable
        SQLCm5.CommandText = "select * from Kenkou order by Kou"
        Adapter5.Fill(Table5)
        DataGridView5.DataSource = Table5

        For i As Integer = 1 To 15
            For r As Integer = 0 To DataGridView5.Rows.Count - 1
                If DataGridView5(0, r).Value = 0 Then
                    CType(Controls("cmbSaiketu" & i), ComboBox).Items.Add(DataGridView5(1, r).Value)
                End If
            Next
        Next

        For i As Integer = 1 To 6
            For r As Integer = 0 To DataGridView5.Rows.Count - 1
                If DataGridView5(0, r).Value = 1 Then
                    CType(Controls("cmbKennsa" & i), ComboBox).Items.Add(DataGridView5(1, r).Value)
                End If
            Next
        Next

        lblHeya.Text = ""
        lblName.Text = ""
        lblHurigana.Text = ""
        YmdBox1.setADStr(Today.ToString("yyyy/MM/dd"))
    End Sub

    Private Function NullCheck(cellvalue As Object) As String
        Return If(IsDBNull(cellvalue), "", cellvalue)
    End Function

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        btnKuria.PerformClick()
        DataGridView4.Columns.Clear()

        Dim s As Integer = DataGridView1.CurrentRow.Index
        lblName.Text = nullcheck(DataGridView1(1, s).Value)
        Dim DGV3rowcount As Integer = DataGridView3.Rows.Count
        If NullCheck(DataGridView1(0, s).Value) <> "" Then         '部屋番号がある行を選択したとき
            If s < 11 Then
                lblHeya.Text = "空　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(2, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(1, i).Value)
                        lblHurigana.Text = NullCheck(DataGridView3(3, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 22 Then
                lblHeya.Text = "森　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(5, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(4, i).Value)
                        lblHurigana.Text = NullCheck(DataGridView3(6, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 33 Then
                lblHeya.Text = "星　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(8, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(7, i).Value)
                        lblHurigana.Text = NullCheck(DataGridView3(9, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 44 Then
                lblHeya.Text = "月　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(11, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(10, i).Value)
                        lblHurigana.Text = NullCheck(DataGridView3(12, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 55 Then
                lblHeya.Text = "花　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(14, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(13, i).Value)
                        lblHurigana.Text = NullCheck(DataGridView3(15, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 66 Then
                lblHeya.Text = "丘　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(17, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(16, i).Value)
                        lblHurigana.Text = NullCheck(DataGridView3(18, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 77 Then
                lblHeya.Text = "虹　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(20, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(19, i).Value)
                        lblHurigana.Text = NullCheck(DataGridView3(21, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 88 Then
                lblHeya.Text = "光　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(23, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(22, i).Value)
                        lblHurigana.Text = NullCheck(DataGridView3(24, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 99 Then
                lblHeya.Text = "雪　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(26, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(25, i).Value)
                        lblHurigana.Text = NullCheck(DataGridView3(27, i).Value)
                        Exit For
                    End If
                Next
            ElseIf s < 110 Then
                lblHeya.Text = "風　" & NullCheck(DataGridView1(0, s).Value)
                For i As Integer = 0 To DGV3rowcount - 1
                    If lblName.Text = NullCheck(DataGridView3(29, i).Value) Then
                        lblID.Text = NullCheck(DataGridView3(28, i).Value)
                        lblHurigana.Text = NullCheck(DataGridView3(30, i).Value)
                        Exit For
                    End If
                Next
            End If
        Else         '部屋番号がない行を選択したとき
            lblHeya.Text = ""
            lblID.Text = ""
            lblName.Text = ""
            lblHurigana.Text = ""
            '何もしない
        End If

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        If lblID.Text <> "" Then
            Dim Table As New DataTable
            SQLCm.CommandText = "select * from Kensin WHERE Id = " & lblID.Text & "AND Nam = '" & lblName.Text & "' order by Gyo"
            Adapter.Fill(Table)
            DataGridView4.DataSource = Table

            'データがあったら表示
            If DataGridView4.Rows.Count > 14 Then
                For Gyo As Integer = 1 To 15
                    Controls("cmbSaiketu" & Gyo).Text = NullCheck(DataGridView4(7, Gyo - 1).Value)
                    If Gyo < 7 Then
                        If IsDBNull(NullCheck(DataGridView4(8, Gyo - 1).Value)) = False Then
                            Controls("cmbKennsa" & Gyo).Text = NullCheck(DataGridView4(8, Gyo - 1).Value)
                        End If
                    End If
                    If Gyo < 5 Then
                        If IsDBNull(NullCheck(DataGridView4(9, Gyo - 1).Value)) = False Then
                            Controls("txtTokki" & Gyo).Text = NullCheck(DataGridView4(9, Gyo - 1).Value)
                        End If
                    End If
                Next
                cmbKikann.Text = NullCheck(DataGridView4(6, 0).Value)
                YmdBox1.setADStr(NullCheck(DataGridView4(10, 0).Value))
            End If
        End If

        If DataGridView2.Rows.Count > 1 Then

        Else
            TopForm.EnableDoubleBuffering(DataGridView2)
            With DataGridView2
                .RowTemplate.Height = 16
                .AllowUserToAddRows = False '行追加禁止
                .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
                .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
                .AllowUserToDeleteRows = False
                .ColumnHeadersVisible = False '列ヘッダー削除
                .ReadOnly = True '編集禁止
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
                .MultiSelect = False
                .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            End With

            Dim SQLCm2 As OleDbCommand = Cn.CreateCommand
            Dim Adapter2 As New OleDbDataAdapter(SQLCm2)
            Dim Table2 As New DataTable
            SQLCm2.CommandText = "select * from Kenkou order by Kou"
            Adapter2.Fill(Table2)
            DataGridView2.DataSource = Table2

            With DataGridView2
                .RowHeadersWidth = 20
                .Columns(0).Visible = False
                .Columns(1).Width = 100
            End With
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
        cmbKikann.Text = ""
        For i As Integer = 1 To 15
            Controls("cmbSaiketu" & i).Text = ""
        Next
        For i As Integer = 1 To 6
            Controls("cmbKennsa" & i).Text = ""
        Next
        For i As Integer = 1 To 4
            Controls("txtTokki" & i).Text = ""
        Next
        YmdBox1.clearText()
    End Sub

    Private Sub btnKousinn_Click(sender As System.Object, e As System.EventArgs) Handles btnKousinn.Click
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)

        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        SQLCm.CommandText = "select * from Kensin WHERE Id = " & lblID.Text & " order by Gyo"
        Adapter.Fill(Table)
        DataGridView4.DataSource = Table

        Dim SQLCm2 As OleDbCommand = Cn.CreateCommand
        Dim Adapter2 As New OleDbDataAdapter(SQLCm2)
        Dim Table2 As New DataTable
        SQLCm2.CommandText = "select * from Kenkou order by Kou"
        Adapter2.Fill(Table2)
        DataGridView2.DataSource = Table2

        For i As Integer = 1 To 15
            CType(Controls("cmbSaiketu" & i), ComboBox).Items.Clear()
        Next

        For i As Integer = 1 To 6
            CType(Controls("cmbKennsa" & i), ComboBox).Items.Clear()
        Next

        For i As Integer = 1 To 15
            For r As Integer = 0 To DataGridView2.Rows.Count - 1
                If DataGridView2(0, r).Value = 0 Then
                    CType(Controls("cmbSaiketu" & i), ComboBox).Items.Add(DataGridView2(1, r).Value)
                End If
            Next
        Next

        For i As Integer = 1 To 6
            For r As Integer = 0 To DataGridView2.Rows.Count - 1
                If DataGridView2(0, r).Value = 1 Then
                    CType(Controls("cmbKennsa" & i), ComboBox).Items.Add(DataGridView2(1, r).Value)
                End If
            Next
        Next
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
        SQL = "DELETE FROM Kensin WHERE Id = " & lblID.Text
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
        For Gyo As Integer = 1 To 15
            If Gyo < 7 Then
                Row("Ken") = Controls("CmbKennsa" & Gyo).Text
            Else
                Row("Ken") = ""
            End If
            If Gyo < 5 Then
                Row("Tokki") = Controls("txtTokki" & Gyo).Text
            Else
                Row("Tokki") = ""
            End If
            Row("Id") = lblID.Text
            Row("Nam") = lblName.Text
            Row("Kana") = lblHurigana.Text
            Row("Rm") = Strings.Right(lblHeya.Text, 3)
            Row("Unt") = Strings.Left(lblHeya.Text, 1)
            Row("Gyo") = Gyo
            Row("Kikan") = cmbKikann.Text
            Row("Sai") = Controls("cmbSaiketu" & Gyo).Text
            Row("Ymd") = YmdBox1.getADStr()
            Dim SQL As String = ""
            SQL = "INSERT INTO Kensin (Id, Nam, Kana, Rm, Unt, Gyo, Kikan, Sai, Ken, Tokki, Ymd) VALUES ("
            SQL &= Row("Id") & ", "
            SQL &= "'" & Row("Nam") & "', "
            SQL &= "'" & Row("Kana") & "', "
            SQL &= "'" & Row("Rm") & "', "
            SQL &= "'" & Row("Unt") & "', "
            SQL &= "'" & Row("Gyo") & "', "
            SQL &= "'" & Row("Kikan") & "', "
            SQL &= "'" & Row("Sai") & "', "
            SQL &= "'" & Row("Ken") & "', "
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
        '採血項目のcmbboxを見る
        For Saiketu As Integer = 1 To 15
            If Controls("cmbSaiketu" & Saiketu).Text <> "" Then     'cmbboxが空じゃないとき
                For i As Integer = 0 To DataGridView2.Rows.Count - 1
                    If DataGridView2(0, i).Value = 0 AndAlso Controls("cmbSaiketu" & Saiketu).Text = DataGridView2(1, i).Value Then  '項目名が一緒
                        GoTo line1  '同じものが見つかったら次のcmbboxを見る
                    End If
                Next
                SaiketuKoumokuTuika(Saiketu)
            Else        'cmbboxが空のとき
                'なにもしない
            End If
line1:
        Next

        For Kensa As Integer = 1 To 6
            If Controls("cmbKennsa" & Kensa).Text <> "" Then     'cmbboxが空じゃないとき
                For i As Integer = 0 To DataGridView2.Rows.Count - 1
                    If DataGridView2(0, i).Value = 1 AndAlso Controls("cmbKennsa" & Kensa).Text = DataGridView2(1, i).Value Then
                        GoTo line2  '同じものが見つかったら次のcmbboxを見る
                    End If
                Next
                KennsaTuika(Kensa)
            Else        'cmbboxが空のとき
                'なにもしない
            End If
line2:
        Next
    End Sub
    Private Sub SaiketuKoumokuTuika(saiketuNo As Integer)
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Table As DataTable = DirectCast(DataGridView2.DataSource, DataTable)
        Dim Row As DataRow = Table.NewRow

        Cn.Open()

        Row("Syu") = 0
        Row("Kou") = Controls("cmbSaiketu" & saiketuNo).Text
        Dim SQL As String = ""
        SQL = "INSERT INTO Kenkou (Syu, Kou) VALUES ("
        SQL &= Row("Syu") & ", "
        SQL &= "'" & Row("Kou") & "' "
        SQL &= ")"
        SQLCm.CommandText = SQL
        SQLCm.ExecuteNonQuery()

        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()

    End Sub
    Private Sub KennsaTuika(kensaNo As Integer)
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Table As DataTable = DirectCast(DataGridView2.DataSource, DataTable)
        Dim Row As DataRow = Table.NewRow

        Cn.Open()

        Row("Syu") = 1
        Row("Kou") = Controls("cmbKennsa" & kensaNo).Text
        Dim SQL As String = ""
        SQL = "INSERT INTO Kenkou (Syu, Kou) VALUES ("
        SQL &= Row("Syu") & ", "
        SQL &= "'" & Row("Kou") & "' "
        SQL &= ")"
        SQLCm.CommandText = SQL
        SQLCm.ExecuteNonQuery()

        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()
    End Sub

    Private Sub btnSakujo_Click(sender As System.Object, e As System.EventArgs) Handles btnSakujo.Click
        Dim i As Integer = DataGridView2.CurrentRow.Index
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim SQL As String = ""
        SQL = "DELETE FROM Kenkou WHERE Syu = '" & DataGridView2(0, i).Value & "' AND Kou = '" & DataGridView2(1, i).Value & "'"
        SQLCm.CommandText = SQL
        Cn.Open()
        SQLCm.ExecuteNonQuery()
        Cn.Close()

        SQLCm.Dispose()
        Cn.Dispose()

        btnKousinn.PerformClick()
    End Sub

    Private Function Nenndo(yyyy As String, MM As String)
        If Val(MM) < 4 Then
            yyyy = Val(yyyy) - 1
            Return yyyy
        Else
            Return yyyy
        End If
    End Function

    Private Sub btnInnsatu_Click(sender As System.Object, e As System.EventArgs) Handles btnInnsatu.Click
        Dim passForm As Form = New passwordForm(TopForm.iniFilePath, 3)
        If passForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        ProgressBar1.Visible = True
        ProgressBar1.Maximum = 10
        ProgressBar1.Minimum = 0
        Label13.Visible = True

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
        oSheet = objWorkBook.Worksheets("健康診断新")

        Dim a As DateTime = Today
        Dim M As String = Strings.Mid(a, 6, 2)
        Dim Nendo As String = Nenndo(Strings.Left(a, 4), M)
        For i As Integer = 2 To 596 Step 66
            oSheet.Range("E" & i).Value = Nendo & "年度　健康診断予定表＜特別養護老人ホーム＞"
            oSheet.Range("Q" & i).Value = Val(M) & "月"
            oSheet.Range("Q" & i + 64).Value = "印刷日：" & a
        Next

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim id As Integer
        Dim DGV3rowcount As Integer = DataGridView3.Rows.Count
        Dim c As Integer = 7

        For count As Integer = 0 To DGV3rowcount - 1    'ユニット居室のデータベース
            For name As Integer = c To 61 Step 6
                '空の家
                oSheet.Range("D" & name).Value = DataGridView3(2, count).Value
                oSheet.Range("D" & name - 1).Value = DataGridView3(3, count).Value

                If Util.checkDBNullValue(DataGridView3(2, count).Value) = "" Then
                    id = 0
                Else
                    id = DataGridView3(1, count).Value
                End If

                Dim TableSora As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from Kensin WHERE Id = " & id & "AND Nam = '" & DataGridView3(2, count).Value & "' order by Gyo"
                Adapter.Fill(TableSora)
                DataGridView6.DataSource = TableSora

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If NullCheck(DataGridView6(5, r).Value) = 1 Then
                        oSheet.Range("I" & name - 2).Value = NullCheck(DataGridView6(6, r).Value)
                        oSheet.Range("G" & name - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name - 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 2 Then
                        oSheet.Range("G" & name).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name - 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 3 Then
                        oSheet.Range("G" & name + 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 4 Then
                        oSheet.Range("G" & name + 2).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 5 Then
                        oSheet.Range("G" & name + 3).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 6 Then
                        oSheet.Range("I" & name - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + 3).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 7 Then
                        oSheet.Range("I" & name).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 8 Then
                        oSheet.Range("I" & name + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 9 Then
                        oSheet.Range("I" & name + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 10 Then
                        oSheet.Range("I" & name + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 11 Then
                        oSheet.Range("K" & name - 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 12 Then
                        oSheet.Range("K" & name).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 13 Then
                        oSheet.Range("K" & name + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 14 Then
                        oSheet.Range("K" & name + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 15 Then
                        oSheet.Range("K" & name + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    End If
                Next

                '森の家
                oSheet.Range("D" & name + (66 * 1)).Value = NullCheck(DataGridView3(5, count).Value)
                oSheet.Range("D" & name + (66 * 1 - 1)).Value = NullCheck(DataGridView3(6, count).Value)

                If NullCheck(DataGridView3(5, count).Value) = "" Then
                    id = 0
                Else
                    id = NullCheck(DataGridView3(4, count).Value)
                End If

                Dim TableMori As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from Kensin WHERE Id = " & id & "AND Nam = '" & NullCheck(DataGridView3(5, count).Value) & "' order by Gyo"
                Adapter.Fill(TableMori)
                DataGridView6.DataSource = TableMori

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If NullCheck(DataGridView6(5, r).Value) = 1 Then
                        oSheet.Range("I" & name + (66 * 1) - 2).Value = NullCheck(DataGridView6(6, r).Value)
                        oSheet.Range("G" & name + (66 * 1) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 1) - 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 2 Then
                        oSheet.Range("G" & name + (66 * 1)).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 1) - 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 3 Then
                        oSheet.Range("G" & name + (66 * 1) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 1)).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 4 Then
                        oSheet.Range("G" & name + (66 * 1) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 1) + 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 5 Then
                        oSheet.Range("G" & name + (66 * 1) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 1) + 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 6 Then
                        oSheet.Range("I" & name + (66 * 1) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 1) + 3).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 7 Then
                        oSheet.Range("I" & name + (66 * 1)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 8 Then
                        oSheet.Range("I" & name + (66 * 1) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 9 Then
                        oSheet.Range("I" & name + (66 * 1) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 10 Then
                        oSheet.Range("I" & name + (66 * 1) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 11 Then
                        oSheet.Range("K" & name + (66 * 1) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 12 Then
                        oSheet.Range("K" & name + (66 * 1)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 13 Then
                        oSheet.Range("K" & name + (66 * 1) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 14 Then
                        oSheet.Range("K" & name + (66 * 1) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 15 Then
                        oSheet.Range("K" & name + (66 * 1) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    End If
                Next

                '星の家
                oSheet.Range("D" & name + (66 * 2)).Value = NullCheck(DataGridView3(8, count).Value)
                oSheet.Range("D" & name + (66 * 2 - 1)).Value = NullCheck(DataGridView3(9, count).Value)

                If NullCheck(DataGridView3(8, count).Value) = "" Then
                    id = 0
                Else
                    id = NullCheck(DataGridView3(7, count).Value)
                End If

                Dim TableHosi As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from Kensin WHERE Id = " & id & "AND Nam = '" & NullCheck(DataGridView3(8, count).Value) & "' order by Gyo"
                Adapter.Fill(TableHosi)
                DataGridView6.DataSource = TableHosi

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If NullCheck(DataGridView6(5, r).Value) = 1 Then
                        oSheet.Range("I" & name + (66 * 2) - 2).Value = NullCheck(DataGridView6(6, r).Value)
                        oSheet.Range("G" & name + (66 * 2) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 2) - 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 2 Then
                        oSheet.Range("G" & name + (66 * 2)).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 2) - 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 3 Then
                        oSheet.Range("G" & name + (66 * 2) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 2)).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 4 Then
                        oSheet.Range("G" & name + (66 * 2) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 2) + 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 5 Then
                        oSheet.Range("G" & name + (66 * 2) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 2) + 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 6 Then
                        oSheet.Range("I" & name + (66 * 2) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 2) + 3).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 7 Then
                        oSheet.Range("I" & name + (66 * 2)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 8 Then
                        oSheet.Range("I" & name + (66 * 2) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 9 Then
                        oSheet.Range("I" & name + (66 * 2) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 10 Then
                        oSheet.Range("I" & name + (66 * 2) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 11 Then
                        oSheet.Range("K" & name + (66 * 2) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 12 Then
                        oSheet.Range("K" & name + (66 * 2)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 13 Then
                        oSheet.Range("K" & name + (66 * 2) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 14 Then
                        oSheet.Range("K" & name + (66 * 2) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 15 Then
                        oSheet.Range("K" & name + (66 * 2) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    End If
                Next

                '月の家
                oSheet.Range("D" & name + (66 * 3)).Value = NullCheck(DataGridView3(11, count).Value)
                oSheet.Range("D" & name + (66 * 3 - 1)).Value = NullCheck(DataGridView3(12, count).Value)

                If NullCheck(DataGridView3(11, count).Value) = "" Then
                    id = 0
                Else
                    id = NullCheck(DataGridView3(10, count).Value)
                End If

                Dim Tabletuki As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from Kensin WHERE Id = " & id & "AND Nam = '" & NullCheck(DataGridView3(11, count).Value) & "' order by Gyo"
                Adapter.Fill(Tabletuki)
                DataGridView6.DataSource = Tabletuki

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If NullCheck(DataGridView6(5, r).Value) = 1 Then
                        oSheet.Range("I" & name + (66 * 3) - 2).Value = NullCheck(DataGridView6(6, r).Value)
                        oSheet.Range("G" & name + (66 * 3) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 3) - 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 2 Then
                        oSheet.Range("G" & name + (66 * 3)).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 3) - 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 3 Then
                        oSheet.Range("G" & name + (66 * 3) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 3)).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 4 Then
                        oSheet.Range("G" & name + (66 * 3) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 3) + 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 5 Then
                        oSheet.Range("G" & name + (66 * 3) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 3) + 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 6 Then
                        oSheet.Range("I" & name + (66 * 3) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 3) + 3).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 7 Then
                        oSheet.Range("I" & name + (66 * 3)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 8 Then
                        oSheet.Range("I" & name + (66 * 3) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 9 Then
                        oSheet.Range("I" & name + (66 * 3) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 10 Then
                        oSheet.Range("I" & name + (66 * 3) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 11 Then
                        oSheet.Range("K" & name + (66 * 3) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 12 Then
                        oSheet.Range("K" & name + (66 * 3)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 13 Then
                        oSheet.Range("K" & name + (66 * 3) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 14 Then
                        oSheet.Range("K" & name + (66 * 3) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 15 Then
                        oSheet.Range("K" & name + (66 * 3) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    End If
                Next

                '花の家
                oSheet.Range("D" & name + (66 * 4)).Value = NullCheck(DataGridView3(14, count).Value)
                oSheet.Range("D" & name + (66 * 4 - 1)).Value = NullCheck(DataGridView3(15, count).Value)

                If NullCheck(DataGridView3(14, count).Value) = "" Then
                    id = 0
                Else
                    id = NullCheck(DataGridView3(13, count).Value)
                End If

                Dim TableHana As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from Kensin WHERE Id = " & id & "AND Nam = '" & NullCheck(DataGridView3(14, count).Value) & "' order by Gyo"
                Adapter.Fill(TableHana)
                DataGridView6.DataSource = TableHana

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If NullCheck(DataGridView6(5, r).Value) = 1 Then
                        oSheet.Range("I" & name + (66 * 4) - 2).Value = NullCheck(DataGridView6(6, r).Value)
                        oSheet.Range("G" & name + (66 * 4) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 4) - 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 2 Then
                        oSheet.Range("G" & name + (66 * 4)).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 4) - 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 3 Then
                        oSheet.Range("G" & name + (66 * 4) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 4)).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 4 Then
                        oSheet.Range("G" & name + (66 * 4) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 4) + 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 5 Then
                        oSheet.Range("G" & name + (66 * 4) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 4) + 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 6 Then
                        oSheet.Range("I" & name + (66 * 4) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 4) + 3).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 7 Then
                        oSheet.Range("I" & name + (66 * 4)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 8 Then
                        oSheet.Range("I" & name + (66 * 4) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 9 Then
                        oSheet.Range("I" & name + (66 * 4) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 10 Then
                        oSheet.Range("I" & name + (66 * 4) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 11 Then
                        oSheet.Range("K" & name + (66 * 4) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 12 Then
                        oSheet.Range("K" & name + (66 * 4)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 13 Then
                        oSheet.Range("K" & name + (66 * 4) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 14 Then
                        oSheet.Range("K" & name + (66 * 4) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 15 Then
                        oSheet.Range("K" & name + (66 * 4) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    End If
                Next

                '丘の家
                oSheet.Range("D" & name + (66 * 5)).Value = NullCheck(DataGridView3(17, count).Value)
                oSheet.Range("D" & name + (66 * 5 - 1)).Value = NullCheck(DataGridView3(18, count).Value)

                If NullCheck(DataGridView3(17, count).Value) = "" Then
                    id = 0
                Else
                    id = NullCheck(DataGridView3(16, count).Value)
                End If

                Dim TableOka As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from Kensin WHERE Id = " & id & "AND Nam = '" & NullCheck(DataGridView3(17, count).Value) & "' order by Gyo"
                Adapter.Fill(TableOka)
                DataGridView6.DataSource = TableOka

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If NullCheck(DataGridView6(5, r).Value) = 1 Then
                        oSheet.Range("I" & name + (66 * 5) - 2).Value = NullCheck(DataGridView6(6, r).Value)
                        oSheet.Range("G" & name + (66 * 5) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 5) - 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 2 Then
                        oSheet.Range("G" & name + (66 * 5)).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 5) - 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 3 Then
                        oSheet.Range("G" & name + (66 * 5) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 5)).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 4 Then
                        oSheet.Range("G" & name + (66 * 5) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 5) + 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 5 Then
                        oSheet.Range("G" & name + (66 * 5) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 5) + 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 6 Then
                        oSheet.Range("I" & name + (66 * 5) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 5) + 3).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 7 Then
                        oSheet.Range("I" & name + (66 * 5)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 8 Then
                        oSheet.Range("I" & name + (66 * 5) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 9 Then
                        oSheet.Range("I" & name + (66 * 5) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 10 Then
                        oSheet.Range("I" & name + (66 * 5) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 11 Then
                        oSheet.Range("K" & name + (66 * 5) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 12 Then
                        oSheet.Range("K" & name + (66 * 5)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 13 Then
                        oSheet.Range("K" & name + (66 * 5) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 14 Then
                        oSheet.Range("K" & name + (66 * 5) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 15 Then
                        oSheet.Range("K" & name + (66 * 5) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    End If
                Next

                '虹の家
                oSheet.Range("D" & name + (66 * 6)).Value = NullCheck(DataGridView3(20, count).Value)
                oSheet.Range("D" & name + (66 * 6 - 1)).Value = NullCheck(DataGridView3(21, count).Value)

                If NullCheck(DataGridView3(20, count).Value) = "" Then
                    id = 0
                Else
                    id = NullCheck(DataGridView3(19, count).Value)
                End If

                Dim TableNiji As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from Kensin WHERE Id = " & id & "AND Nam = '" & NullCheck(DataGridView3(20, count).Value) & "' order by Gyo"
                Adapter.Fill(TableNiji)
                DataGridView6.DataSource = TableNiji

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If NullCheck(DataGridView6(5, r).Value) = 1 Then
                        oSheet.Range("I" & name + (66 * 6) - 2).Value = NullCheck(DataGridView6(6, r).Value)
                        oSheet.Range("G" & name + (66 * 6) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 6) - 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 2 Then
                        oSheet.Range("G" & name + (66 * 6)).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 6) - 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 3 Then
                        oSheet.Range("G" & name + (66 * 6) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 6)).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 4 Then
                        oSheet.Range("G" & name + (66 * 6) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 6) + 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 5 Then
                        oSheet.Range("G" & name + (66 * 6) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 6) + 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 6 Then
                        oSheet.Range("I" & name + (66 * 6) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 6) + 3).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 7 Then
                        oSheet.Range("I" & name + (66 * 6)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 8 Then
                        oSheet.Range("I" & name + (66 * 6) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 9 Then
                        oSheet.Range("I" & name + (66 * 6) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 10 Then
                        oSheet.Range("I" & name + (66 * 6) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 11 Then
                        oSheet.Range("K" & name + (66 * 6) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 12 Then
                        oSheet.Range("K" & name + (66 * 6)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 13 Then
                        oSheet.Range("K" & name + (66 * 6) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 14 Then
                        oSheet.Range("K" & name + (66 * 6) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 15 Then
                        oSheet.Range("K" & name + (66 * 6) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    End If
                Next

                '光の家
                oSheet.Range("D" & name + (66 * 7)).Value = NullCheck(DataGridView3(23, count).Value)
                oSheet.Range("D" & name + (66 * 7 - 1)).Value = NullCheck(DataGridView3(24, count).Value)

                If NullCheck(DataGridView3(23, count).Value) = "" Then
                    id = 0
                Else
                    id = NullCheck(DataGridView3(22, count).Value)
                End If

                Dim TableHikari As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from Kensin WHERE Id = " & id & "AND Nam = '" & NullCheck(DataGridView3(23, count).Value) & "' order by Gyo"
                Adapter.Fill(TableHikari)
                DataGridView6.DataSource = TableHikari

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If NullCheck(DataGridView6(5, r).Value) = 1 Then
                        oSheet.Range("I" & name + (66 * 7) - 2).Value = NullCheck(DataGridView6(6, r).Value)
                        oSheet.Range("G" & name + (66 * 7) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 7) - 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 2 Then
                        oSheet.Range("G" & name + (66 * 7)).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 7) - 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 3 Then
                        oSheet.Range("G" & name + (66 * 7) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 7)).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 4 Then
                        oSheet.Range("G" & name + (66 * 7) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 7) + 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 5 Then
                        oSheet.Range("G" & name + (66 * 7) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 7) + 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 6 Then
                        oSheet.Range("I" & name + (66 * 7) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 7) + 3).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 7 Then
                        oSheet.Range("I" & name + (66 * 7)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 8 Then
                        oSheet.Range("I" & name + (66 * 7) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 9 Then
                        oSheet.Range("I" & name + (66 * 7) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 10 Then
                        oSheet.Range("I" & name + (66 * 7) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 11 Then
                        oSheet.Range("K" & name + (66 * 7) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 12 Then
                        oSheet.Range("K" & name + (66 * 7)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 13 Then
                        oSheet.Range("K" & name + (66 * 7) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 14 Then
                        oSheet.Range("K" & name + (66 * 7) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 15 Then
                        oSheet.Range("K" & name + (66 * 7) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    End If
                Next

                '雪の家
                oSheet.Range("D" & name + (66 * 8)).Value = NullCheck(DataGridView3(26, count).Value)
                oSheet.Range("D" & name + (66 * 8 - 1)).Value = NullCheck(DataGridView3(27, count).Value)

                If NullCheck(DataGridView3(26, count).Value) = "" Then
                    id = 0
                Else
                    id = NullCheck(DataGridView3(25, count).Value)
                End If

                Dim TableYuki As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from Kensin WHERE Id = " & id & "AND Nam = '" & NullCheck(DataGridView3(26, count).Value) & "' order by Gyo"
                Adapter.Fill(TableYuki)
                DataGridView6.DataSource = TableYuki

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If NullCheck(DataGridView6(5, r).Value) = 1 Then
                        oSheet.Range("I" & name + (66 * 8) - 2).Value = NullCheck(DataGridView6(6, r).Value)
                        oSheet.Range("G" & name + (66 * 8) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 8) - 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 2 Then
                        oSheet.Range("G" & name + (66 * 8)).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 8) - 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 3 Then
                        oSheet.Range("G" & name + (66 * 8) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 8)).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 4 Then
                        oSheet.Range("G" & name + (66 * 8) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 8) + 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 5 Then
                        oSheet.Range("G" & name + (66 * 8) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 8) + 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 6 Then
                        oSheet.Range("I" & name + (66 * 8) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 8) + 3).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 7 Then
                        oSheet.Range("I" & name + (66 * 8)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 8 Then
                        oSheet.Range("I" & name + (66 * 8) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 9 Then
                        oSheet.Range("I" & name + (66 * 8) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 10 Then
                        oSheet.Range("I" & name + (66 * 8) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 11 Then
                        oSheet.Range("K" & name + (66 * 8) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 12 Then
                        oSheet.Range("K" & name + (66 * 8)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 13 Then
                        oSheet.Range("K" & name + (66 * 8) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 14 Then
                        oSheet.Range("K" & name + (66 * 8) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 15 Then
                        oSheet.Range("K" & name + (66 * 8) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    End If
                Next

                '風の家
                oSheet.Range("D" & name + (66 * 9)).Value = NullCheck(DataGridView3(29, count).Value)
                oSheet.Range("D" & name + (66 * 9 - 1)).Value = NullCheck(DataGridView3(30, count).Value)

                If NullCheck(DataGridView3(29, count).Value) = "" Then
                    id = 0
                Else
                    id = NullCheck(DataGridView3(28, count).Value)
                End If

                Dim TableKaze As New DataTable
                '対象となる人のデータを表示
                SQLCm.CommandText = "select * from Kensin WHERE Id = " & id & "AND Nam = '" & NullCheck(DataGridView3(29, count).Value) & "' order by Gyo"
                Adapter.Fill(TableKaze)
                DataGridView6.DataSource = TableKaze

                '個人データをエクセルに出力
                For r As Integer = 0 To DataGridView6.Rows.Count - 1
                    If NullCheck(DataGridView6(5, r).Value) = 1 Then
                        oSheet.Range("I" & name + (66 * 9) - 2).Value = NullCheck(DataGridView6(6, r).Value)
                        oSheet.Range("G" & name + (66 * 9) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 9) - 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 2 Then
                        oSheet.Range("G" & name + (66 * 9)).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 9) - 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 3 Then
                        oSheet.Range("G" & name + (66 * 9) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 9)).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 4 Then
                        oSheet.Range("G" & name + (66 * 9) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 9) + 1).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 5 Then
                        oSheet.Range("G" & name + (66 * 9) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 9) + 2).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 6 Then
                        oSheet.Range("I" & name + (66 * 9) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                        oSheet.Range("M" & name + (66 * 9) + 3).Value = NullCheck(DataGridView6(8, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 7 Then
                        oSheet.Range("I" & name + (66 * 9)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 8 Then
                        oSheet.Range("I" & name + (66 * 9) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 9 Then
                        oSheet.Range("I" & name + (66 * 9) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 10 Then
                        oSheet.Range("I" & name + (66 * 9) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 11 Then
                        oSheet.Range("K" & name + (66 * 9) - 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 12 Then
                        oSheet.Range("K" & name + (66 * 9)).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 13 Then
                        oSheet.Range("K" & name + (66 * 9) + 1).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 14 Then
                        oSheet.Range("K" & name + (66 * 9) + 2).Value = NullCheck(DataGridView6(7, r).Value)
                    ElseIf NullCheck(DataGridView6(5, r).Value) = 15 Then
                        oSheet.Range("K" & name + (66 * 9) + 3).Value = NullCheck(DataGridView6(7, r).Value)
                    End If
                Next

                c = name + 6
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
        Label13.Visible = False
    End Sub
End Class