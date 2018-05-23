Imports System.Data.OleDb

Public Class 諸マスタ

    'dgvの幅サイズ
    Private Const DGV_WIDTH_MAX As Integer = 209
    Private Const DGV_WIDTH_MIN As Integer = 192

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

    Private Sub 諸マスタ_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        Me.Left = 0
        Me.Top = 55
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        'テキストボックス設定
        settingInputTextBox()

    End Sub

    Private Sub 諸マスタ_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If e.Control = False Then
                Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
            End If
        End If
    End Sub

    Private Sub displayVariousMasterData(item As String)
        dgvVarious.Columns.Clear()
        'dgv表示前設定
        settingDgv()

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        SQLCm.CommandText = "select Naiyo from EtcM where Komoku=@item order by Naiyo"
        SQLCm.Parameters.Clear()
        SQLCm.Parameters.Add("@item", OleDbType.Char).Value = item
        Cn.Open()
        Adapter.Fill(Table)
        dgvVarious.DataSource = Table
        Cn.Close()
        Cn.Dispose()

        'dgv表示後設定
        settingDgvColumns()

        '内容入力ボックスへフォーカス
        contentBox.Focus()
    End Sub

    Private Sub settingDgv()
        TopForm.EnableDoubleBuffering(dgvVarious)

        With dgvVarious
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .ReadOnly = True
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 19
            .RowTemplate.Height = 16
            .RowHeadersWidth = 35
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .RowTemplate.HeaderCell = New dgvRowHeaderCell() '行ヘッダの三角マークを非表示に
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ShowCellToolTips = False
        End With

    End Sub

    Private Sub settingDgvColumns()
        With dgvVarious
            With .Columns("Naiyo")
                .HeaderText = "内容"
                .Width = 155
                .SortMode = DataGridViewColumnSortMode.NotSortable
            End With

            If .Rows.Count <= 20 Then
                .Size = New Size(DGV_WIDTH_MIN, .Size.Height)
            Else
                .Size = New Size(DGV_WIDTH_MAX, .Size.Height)
            End If
        End With
    End Sub

    Private Sub settingInputTextBox()
        '項目
        settingItemBox()
        cmbItem.DropDownStyle = ComboBoxStyle.DropDownList
        cmbItem.FlatStyle = FlatStyle.Flat

        cmbItem.ImeMode = Windows.Forms.ImeMode.Hiragana

        '内容
        contentBox.ImeMode = Windows.Forms.ImeMode.Hiragana
    End Sub

    Private Sub settingItemBox()
        'クリア
        cmbItem.Items.Clear()

        'コンボボックスのリスト取得、設定
        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        SQLCm.CommandText = "select distinct Komoku from EtcM order by Komoku Desc"
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        While reader.Read() = True
            cmbItem.Items.Add(reader("Komoku"))
        End While
        reader.Close()
        Cn.Close()
    End Sub

    Private Function checkDBNullValue(dgvCellValue As Object) As String
        Return If(IsDBNull(dgvCellValue), "", dgvCellValue)
    End Function

    Private Sub dgvVarious_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvVarious.CellMouseClick
        If e.RowIndex = -1 Then
            Return
        End If

        '内容取得、設定
        contentBox.Text = checkDBNullValue(dgvVarious("Naiyo", e.RowIndex).Value)

        '内容入力ボックスへフォーカス
        contentBox.Focus()

        contentBox.SelectionStart = 0
        contentBox.SelectionLength = 0
    End Sub

    Private Sub dgvVarious_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvVarious.CellPainting
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

    Private Sub cmbItem_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbItem.SelectedValueChanged
        displayVariousMasterData(cmbItem.Text)
        contentBox.Text = ""
    End Sub

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        Dim itemName As String = cmbItem.Text
        Dim content As String = contentBox.Text

        If itemName = "" Then
            MsgBox("項目を選択して下さい。", , "登録エラー")
            Return
        ElseIf content = "" Then
            MsgBox("内容を入力して下さい。", , "登録エラー")
            Return
        End If

        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand

        '入力データ存在チェック
        SQLCm.CommandText = "SELECT top 1 * FROM EtcM where Komoku=@item and naiyo=@content"
        SQLCm.Parameters.Clear()
        SQLCm.Parameters.Add("@item", OleDbType.Char).Value = itemName
        SQLCm.Parameters.Add("@content", OleDbType.Char).Value = content
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        If reader.Read() = True Then
            'データが存在する場合、何もしない
            reader.Close()
            Cn.Close()
        Else
            'データが存在しない場合、登録処理
            reader.Close()
            SQLCm.CommandText = "insert into EtcM(Komoku, Naiyo) values (@item, @content)"
            SQLCm.Parameters.Clear()
            SQLCm.Parameters.Add("@item", OleDbType.Char).Value = itemName
            SQLCm.Parameters.Add("@content", OleDbType.Char).Value = content
            SQLCm.ExecuteNonQuery()
            Cn.Close()
        End If

        'クリア
        contentBox.Text = ""

        '再表示
        displayVariousMasterData(cmbItem.Text)
    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Dim itemName As String = cmbItem.Text
        Dim content As String = contentBox.Text

        If itemName = "" Then
            MsgBox("項目を選択して下さい。", , "削除エラー")
            Return
        ElseIf content = "" Then
            MsgBox("内容を入力して下さい。", , "削除エラー")
            Return
        End If

        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand

        '入力データ存在チェック
        SQLCm.CommandText = "SELECT top 1 * FROM EtcM where Komoku=@item and naiyo=@content"
        SQLCm.Parameters.Clear()
        SQLCm.Parameters.Add("@item", OleDbType.Char).Value = itemName
        SQLCm.Parameters.Add("@content", OleDbType.Char).Value = content
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        If reader.Read() = True Then
            'データが存在する場合、削除処理
            reader.Close()
            Dim result As DialogResult = MessageBox.Show("削除してよろしいですか？", "Nurse2", MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.Yes Then
                SQLCm.CommandText = "delete from EtcM where Komoku=@item and Naiyo=@content"
                SQLCm.Parameters.Clear()
                SQLCm.Parameters.Add("@item", OleDbType.Char).Value = itemName
                SQLCm.Parameters.Add("@content", OleDbType.Char).Value = content
                SQLCm.ExecuteNonQuery()
                Cn.Close()

                'クリア
                contentBox.Text = ""

                '再表示
                displayVariousMasterData(cmbItem.Text)
            Else
                Cn.Close()
            End If
        Else
            'データが存在しない場合
            reader.Close()
            Cn.Close()
            MsgBox("登録されていません。", , "削除エラー")
            Return
        End If
    End Sub
End Class