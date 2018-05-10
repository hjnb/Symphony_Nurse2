Imports System.Data.OleDb

Public Class 処置マスタ

    'スクロール位置保持用
    Private scrollPosition As Integer = 0

    '選択行位置保持用
    Private selectedRowIndex As Integer = 0

    '分類１のテキスト変更時のイベント発生制御フラグ
    Private isInitCategory1 As Boolean = False

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
        settingDgv(dgvTreatingMaster)

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        If category1 = "" Then
            SQLCm.CommandText = "select Bunrui1, Dsp, Bunrui2, Bunrui3 from SyotiM order by Bunrui1, Bunrui2"
        Else
            SQLCm.CommandText = "select Bunrui1, Dsp, Bunrui2, Bunrui3 from SyotiM where Bunrui1=@category1 order by Bunrui1, Bunrui2"
            SQLCm.Parameters.Clear()
            SQLCm.Parameters.Add("@category1", OleDbType.Char).Value = category1
        End If

        Cn.Open()
        Adapter.Fill(Table)
        dgv.DataSource = Table
        Cn.Close()
        Cn.Dispose()

        settingDgvColumns(dgvTreatingMaster)

        'スクロールの位置設定
        dgvTreatingMaster.FirstDisplayedScrollingRowIndex = scrollPosition

        '選択行設定
        dgvTreatingMaster.Rows(selectedRowIndex).Selected = True

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

        End With
    End Sub

    Private Sub settingCategory1Box()
        '
        isInitCategory1 = False
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

        isInitCategory1 = True
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

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click

    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click

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

    Private Sub cmbCategory1_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbCategory1.SelectedValueChanged
        If isInitCategory1 = True Then
            MsgBox("aa")
        End If
    End Sub
End Class