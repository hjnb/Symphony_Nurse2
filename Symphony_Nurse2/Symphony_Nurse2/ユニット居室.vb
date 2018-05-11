Imports System.Data.OleDb

Public Class ユニット居室

    '選択している名前（漢字）保持用
    Private selectedNam As String = ""

    '選択している名前（カナ）保持用
    Private selectedKana As String = ""

    '
    Private unitDt As DataTable

    '
    Private unitAdapter As OleDbDataAdapter

    Private Class UnitData
        Private id As String

    End Class


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

    Private Sub ユニット居室_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Left = 0
        Me.Top = 55
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        '利用者データ表示
        displayUserMasterData()

        'ユニット居室データ表示
        displayUnitMasterData()

    End Sub

    Private Sub displayUserMasterData()
        '表示前設定
        settingDgvUser()

        'データ取得、表示
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        SQLCm.CommandText = "select Id, Nam, Kana, Dsp from KihonM where Dsp=1 order by Kana"
        Adapter.Fill(Table)
        dgvUser.DataSource = Table

        '表示後設定
        settingDgvUserColumns()
    End Sub

    Private Sub displayUnitMasterData()
        '表示前設定
        settingDgvUnit()

        '
        'データ取得、表示



        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        unitAdapter = New OleDbDataAdapter(SQLCm)
        unitDt = New DataTable()
        SQLCm.CommandText = "select * from Unit where order by Gyo"


        '表示後設定
        settingDgvUnitColumns()
    End Sub

    Private Sub settingDgvUser()
        'DoubleBufferedプロパティをTrue
        TopForm.EnableDoubleBuffering(dgvUser)

        'dgv設定
        With dgvUser
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .ReadOnly = True '編集禁止
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect 'クリック時に行選択
            .MultiSelect = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .RowTemplate.Height = 15
            .RowHeadersWidth = 35
            .ColumnHeadersHeight = 20
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .RowTemplate.HeaderCell = New dgvRowHeaderCell() '行ヘッダの三角マークを非表示に
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ShowCellToolTips = False
        End With

    End Sub

    Private Sub settingDgvUserColumns()
        '並び替えができないようにする
        For Each c As DataGridViewColumn In dgvUser.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        With dgvUser
            'ID
            With .Columns("Id")
                .Width = 60
                .HeaderText = "ID"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            '利用者氏名
            With .Columns("Nam")
                .Width = 100
                .HeaderText = "利用者氏名"
            End With

            '非表示
            .Columns("Kana").Visible = False
            .Columns("Dsp").Visible = False

        End With
    End Sub

    Private Sub settingDgvUnit()
        'DoubleBufferedプロパティをTrue
        TopForm.EnableDoubleBuffering(dgvUnit)
    End Sub

    Private Sub settingDgvUnitColumns()

    End Sub

    Private Function checkDBNullValue(dgvCellValue As Object) As String
        Return If(IsDBNull(dgvCellValue), "", dgvCellValue)
    End Function

    Private Sub dgvUser_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvUser.CellMouseClick
        If e.RowIndex = -1 Then
            Return
        End If

        '選択している名前（漢字、カナ）を保持
        selectedNam = checkDBNullValue(dgvUser("Nam", e.RowIndex).Value)
        selectedKana = checkDBNullValue(dgvUser("Kana", e.RowIndex).Value)
    End Sub

    Private Sub dgvUser_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvUser.CellPainting
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

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click

    End Sub
End Class