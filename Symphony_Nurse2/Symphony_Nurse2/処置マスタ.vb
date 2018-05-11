Imports System.Data.OleDb

Public Class 処置マスタ

    '分類１のテキスト変更時のイベント発生制御フラグ
    Private canChangeCategory1Event As Boolean = True

    'テキストボックスのマウスダウンイベント制御用
    Private mdFlag As Boolean = False

    'dgvの幅サイズ
    Private Const DGV_WIDTH_MAX As Integer = 469
    Private Const DGV_WIDTH_MIN As Integer = 452

    '行ヘッダーのカレントセルを表す三角マークを非表示に設定する為のクラス。
    Public Class dgvRowHeaderCell

        'DataGridViewRowHeaderCell を継承
        Inherits DataGridViewRowHeaderCell

        'DataGridViewHeaderCell.Paint をオーバーライドして行ヘッダーを描画
        Protected Overrides Sub Paint(ByVal graphics As Graphics, ByVal clipBounds As Rectangle, _
           ByVal cellBounds As Rectangle, ByVal rowIndex As Integer, ByVal cellState As DataGridViewElementStates, _
           ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, _
           ByVal cellStyle As DataGridViewCellStyle, ByVal advancedBorderStyle As DataGridViewAdvancedBorderStyle, _
           ByVal paintParts As DataGridViewPaintParts)
            '標準セルの描画からセル内容の背景だけ除いた物を描画(-5)
            MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, _
                     formattedValue, errorText, cellStyle, advancedBorderStyle, _
                     Not DataGridViewPaintParts.ContentBackground)
        End Sub

    End Class

    Private Sub 処置マスタ_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        Me.Left = 0
        Me.Top = 55
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        'テキストボックス設定
        settingInputTextBox()

        'dgvにデータ表示
        displayTreatingMasterData(dgvTreatingMaster, cmbCategory1.Text)

    End Sub

    Private Sub 処置マスタ_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If e.Control = False Then
                Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
            End If
        End If
    End Sub

    Private Sub displayTreatingMasterData(dgv As DataGridView, category1 As String)
        dgv.Columns.Clear()
        'dgv表示前設定
        settingDgv(dgvTreatingMaster)

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        If category1 = "" Then
            SQLCm.CommandText = "select Bunrui1, Dsp, Bunrui2, Bunrui3 from SyotiM order by Bunrui1, Dsp, Bunrui2, Bunrui3"
        Else
            SQLCm.CommandText = "select Bunrui1, Dsp, Bunrui2, Bunrui3 from SyotiM where Bunrui1=@category1 order by Dsp, Bunrui2, Bunrui3"
            SQLCm.Parameters.Clear()
            SQLCm.Parameters.Add("@category1", OleDbType.Char).Value = category1
        End If

        Cn.Open()
        Adapter.Fill(Table)
        dgv.DataSource = Table
        Cn.Close()
        Cn.Dispose()

        'dgv表示後設定
        settingDgvColumns(dgvTreatingMaster)

    End Sub

    Private Sub settingDgv(dgv As DataGridView)
        TopForm.EnableDoubleBuffering(dgv)

        With dgv
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .ReadOnly = True
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 20
            .RowTemplate.Height = 15
            .RowHeadersWidth = 35
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .RowTemplate.HeaderCell = New dgvRowHeaderCell() '行ヘッダの三角マークを非表示に
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ShowCellToolTips = False
        End With

    End Sub

    Private Sub settingDgvColumns(dgv As DataGridView)
        With dgv
            '並び替えができないようにする
            For Each c As DataGridViewColumn In dgv.Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            With .Columns("Bunrui1")
                .HeaderText = "分類１"
                .Width = 85
            End With

            With .Columns("Dsp")
                .HeaderText = "表示順"
                .Width = 55
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            With .Columns("Bunrui2")
                .HeaderText = "分類２"
                .Width = 220
            End With

            With .Columns("Bunrui3")
                .HeaderText = "分類３"
                .Width = 55
            End With

            If .Rows.Count <= 20 Then
                .Size = New Size(DGV_WIDTH_MIN, .Size.Height)
            Else
                .Size = New Size(DGV_WIDTH_MAX, .Size.Height)
            End If

        End With
    End Sub

    Private Sub settingCategory1Box()
        'クリア
        cmbCategory1.Items.Clear()

        'コンボボックスのリスト取得、設定
        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        SQLCm.CommandText = "select distinct Bunrui1 from SyotiM order by Bunrui1"
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        While reader.Read() = True
            cmbCategory1.Items.Add(reader("Bunrui1"))
        End While
        reader.Close()
        Cn.Close()
    End Sub

    Private Sub settingInputTextBox()
        '分類１
        settingCategory1Box()
        cmbCategory1.ImeMode = Windows.Forms.ImeMode.Hiragana

        '表示順
        dspBox.ImeMode = Windows.Forms.ImeMode.Disable
        dspBox.TextAlign = HorizontalAlignment.Center

        '分類２
        category2Box.ImeMode = Windows.Forms.ImeMode.Hiragana

        '分類３
        category3Box.ImeMode = Windows.Forms.ImeMode.Hiragana

    End Sub

    Private Sub inputClear()
        cmbCategory1.Text = ""
        dspBox.Text = ""
        category2Box.Text = ""
        category3Box.Text = ""
    End Sub

    Private Function checkDBNullValue(dgvCellValue As Object) As String
        Return If(IsDBNull(dgvCellValue), "", dgvCellValue)
    End Function

    Private Sub searchDgvRow(category1 As String, category2 As String, category3 As String)
        For Each row As DataGridViewRow In dgvTreatingMaster.Rows
            If row.Cells("Bunrui1").Value = category1 AndAlso row.Cells("Bunrui2").Value = category2 AndAlso row.Cells("Bunrui3").Value = category3 Then
                '対象の行の選択、スクロールさせる
                row.Selected = True
                dgvTreatingMaster.FirstDisplayedScrollingRowIndex = row.Index
                Exit For
            End If
        Next
    End Sub

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        '分類１、２、３の入力テキスト取得
        Dim category1 As String = cmbCategory1.Text
        Dim category2 As String = category2Box.Text
        Dim category3 As String = category3Box.Text

        '入力チェック
        If category1 = "" Then
            MsgBox("分類１を入力して下さい。", , "登録エラー")
            Return
        ElseIf dspBox.Text = "" Then
            MsgBox("表示順を入力して下さい。", , "登録エラー")
            Return
        ElseIf IsNumeric(dspBox.Text) = False Then
            MsgBox("表示順は数値を入力して下さい。", , "登録エラー")
            Return
        ElseIf category2 = "" Then
            MsgBox("分類２を入力して下さい。", , "登録エラー")
            Return
        End If

        '表示順に入力された数値の絶対値を取得
        Dim dsp As Integer = Math.Abs(CInt(dspBox.Text))

        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand

        '分類１且つ分類２且つ分類３のデータの存在の有無で、更新か新規登録処理
        SQLCm.CommandText = "SELECT top 1 * FROM SyotiM where Bunrui1=@category1 and Bunrui2=@category2 and Bunrui3=@category3"
        SQLCm.Parameters.Clear()
        SQLCm.Parameters.Add("@category1", OleDbType.Char).Value = category1
        SQLCm.Parameters.Add("@category2", OleDbType.Char).Value = category2
        SQLCm.Parameters.Add("@category3", OleDbType.Char).Value = category3
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        If reader.Read() = True Then
            reader.Close()
            '更新
            SQLCm.CommandText = "update SyotiM set Dsp=@dsp where Bunrui1=@category1 and Bunrui2=@category2 and Bunrui3=@category3"
            SQLCm.Parameters.Clear()
            SQLCm.Parameters.Add("@dsp", OleDbType.Integer).Value = dsp
            SQLCm.Parameters.Add("@category1", OleDbType.Char).Value = category1
            SQLCm.Parameters.Add("@category2", OleDbType.Char).Value = category2
            SQLCm.Parameters.Add("@category3", OleDbType.Char).Value = category3
            SQLCm.ExecuteNonQuery()
            Cn.Close()
        Else
            reader.Close()
            '新規登録
            SQLCm.CommandText = "insert into SyotiM(Bunrui1, Dsp, Bunrui2, Bunrui3) values (@category1, @dsp, @category2, @category3)"
            SQLCm.Parameters.Clear()
            SQLCm.Parameters.Add("@category1", OleDbType.Char).Value = category1
            SQLCm.Parameters.Add("@dsp", OleDbType.Integer).Value = dsp
            SQLCm.Parameters.Add("@category2", OleDbType.Char).Value = category2
            SQLCm.Parameters.Add("@category3", OleDbType.Char).Value = category3
            SQLCm.ExecuteNonQuery()
            Cn.Close()
        End If

        '入力テキストクリア
        inputClear()

        'コンボボックスの値再取得、設定(コンボボックス変更イベントでdgv再表示される)
        settingCategory1Box()
        cmbCategory1.Text = category1

        '追加or更新した行を選択、スクロール
        searchDgvRow(category1, category2, category3)

    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Dim category1 As String = cmbCategory1.Text '分類１
        Dim category2 As String = category2Box.Text '分類２
        Dim category3 As String = category3Box.Text '分類３

        '入力チェック
        If category1 = "" Then
            MsgBox("分類１がありません。", , "削除エラー")
            Return
        ElseIf category2 = "" Then
            MsgBox("分類２がありません。", , "削除エラー")
            Return
        End If

        '入力されているデータが存在するかチェック
        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        SQLCm.CommandText = "SELECT top 1 * from SyotiM where Bunrui1=@category1 and Bunrui2=@category2 and Bunrui3=@category3"
        SQLCm.Parameters.Clear()
        SQLCm.Parameters.Add("@category1", OleDbType.Char).Value = category1
        SQLCm.Parameters.Add("@category2", OleDbType.Char).Value = category2
        SQLCm.Parameters.Add("@category3", OleDbType.Char).Value = category3
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        If reader.Read() = False Then
            reader.Close()
            MsgBox("登録されていません。", , "削除エラー")
            Cn.Close()
            Cn.Dispose()
            Return
        Else
            '削除処理
            reader.Close()
            Dim result As DialogResult = MessageBox.Show("削除してよろしいですか？", "Nurse2", MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.Yes Then
                SQLCm.CommandText = "delete from SyotiM where Bunrui1=@category1 and Bunrui2=@category2 and Bunrui3=@category3"
                SQLCm.Parameters.Clear()
                SQLCm.Parameters.Add("@category1", OleDbType.Char).Value = category1
                SQLCm.Parameters.Add("@category2", OleDbType.Char).Value = category2
                SQLCm.Parameters.Add("@category3", OleDbType.Char).Value = category3
                SQLCm.ExecuteNonQuery()
                Cn.Close()

                '入力テキストクリア
                inputClear()

                'コンボボックスの値再取得、設定(コンボボックス変更イベントでdgv再表示される)
                settingCategory1Box()
                If cmbCategory1.FindString(category1) = -1 Then
                    displayTreatingMasterData(dgvTreatingMaster, "")
                Else
                    cmbCategory1.Text = category1
                End If
            End If
        End If
    End Sub

    Private Sub dgvTreatingMaster_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvTreatingMaster.CellMouseClick
        If e.RowIndex = -1 Then
            Return
        End If

        '選択セルから値取得
        Dim category1 As String = checkDBNullValue(dgvTreatingMaster("Bunrui1", e.RowIndex).Value)
        Dim dsp As String = checkDBNullValue(dgvTreatingMaster("Dsp", e.RowIndex).Value)
        Dim category2 As String = checkDBNullValue(dgvTreatingMaster("Bunrui2", e.RowIndex).Value)
        Dim category3 As String = checkDBNullValue(dgvTreatingMaster("Bunrui3", e.RowIndex).Value)

        '分類１イベント制御をfalseに
        canChangeCategory1Event = False

        'テキストボックスへセット
        cmbCategory1.Text = category1
        dspBox.Text = dsp
        dspBox.Focus()
        category2Box.Text = category2
        category3Box.Text = category3

        '分類１イベント制御をtrueに
        canChangeCategory1Event = True
    End Sub

    Private Sub dgvTreatingMaster_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvTreatingMaster.CellPainting
        '行ヘッダーかどうか調べる
        If e.ColumnIndex < 0 AndAlso e.RowIndex >= 0 Then
            'セルを描画する
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All)

            '行番号を描画する範囲を決定する
            'e.AdvancedBorderStyleやe.CellStyle.Paddingは無視しています
            Dim indexRect As Rectangle = e.CellBounds
            indexRect.Inflate(-2, -2)
            '行番号を描画する
            TextRenderer.DrawText(e.Graphics, _
                (e.RowIndex + 1).ToString(), _
                e.CellStyle.Font, _
                indexRect, _
                e.CellStyle.ForeColor, _
                TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            '描画が完了したことを知らせる
            e.Handled = True
        End If
    End Sub

    Private Sub cmbCategory1_Click(sender As Object, e As System.EventArgs) Handles cmbCategory1.Click
        inputClear()
    End Sub

    Private Sub cmbCategory1_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbCategory1.SelectedValueChanged
        If canChangeCategory1Event = True Then
            displayTreatingMasterData(dgvTreatingMaster, cmbCategory1.Text)
        End If
    End Sub

    Private Sub textBox_Enter(sender As Object, e As System.EventArgs) Handles dspBox.Enter, category2Box.Enter, category3Box.Enter
        Dim tb As TextBox = CType(sender, TextBox)
        tb.SelectAll()
        mdFlag = True
    End Sub

    Private Sub textBox_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dspBox.MouseDown, category2Box.MouseDown, category3Box.MouseDown
        If mdFlag = True Then
            Dim tb As TextBox = CType(sender, TextBox)
            tb.SelectAll()
            mdFlag = False
        End If
    End Sub
End Class