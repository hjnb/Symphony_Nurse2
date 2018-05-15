Imports System.Data.OleDb

Public Class 体重管理

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

    '選択している名前保持用
    Private selectedNam As String = ""

    'ラジオボタン背景色
    Private rbtnBackColor As Color = Color.FromArgb(255, 255, 0)

    '名前表示色
    Private displayNameColor As Color = Color.FromArgb(0, 120, 215)

    Private Sub 体重管理_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Left = 0
        Me.Top = 55
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        rbtnHonkan.Checked = True
        selectUserLabel.ForeColor = displayNameColor
        initYmdBox()

        displayUserMasterData()
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

    Private Sub settingUnitLabel(unitType As String)
        If unitType = "本館" Then
            labelUnit1.Text = "空の家"
            labelUnit2.Text = "森の家"
            labelUnit3.Text = "星の家"
            labelUnit4.Text = "月の家"
            labelUnit5.Text = "花の家"
            labelUnit6.Text = "海の家"
            labelUnit6.Visible = True
            ymdBoxUnit6.Visible = True
        ElseIf unitType = "ｱﾈｯｸｽ" Then
            labelUnit1.Text = "丘の家"
            labelUnit2.Text = "虹の家"
            labelUnit3.Text = "光の家"
            labelUnit4.Text = "雪の家"
            labelUnit5.Text = "風の家"
            labelUnit6.Visible = False
            ymdBoxUnit6.Visible = False
        End If
    End Sub

    Private Sub initYmdBox()
        '日付を当日で設定
        Dim nowStr As String = DateTime.Now.ToString("yyyy/MM/dd")
        ymdBoxUnit1.setADStr(nowStr)
        ymdBoxUnit2.setADStr(nowStr)
        ymdBoxUnit3.setADStr(nowStr)
        ymdBoxUnit4.setADStr(nowStr)
        ymdBoxUnit5.setADStr(nowStr)
        ymdBoxUnit6.setADStr(nowStr)
    End Sub

    Private Function checkDBNullValue(dgvCellValue As Object) As String
        Return If(IsDBNull(dgvCellValue), "", dgvCellValue)
    End Function

    Private Sub dgvUser_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvUser.CellMouseClick
        If e.RowIndex = -1 Then
            Return
        End If

        '選択した名前保持
        selectedNam = checkDBNullValue(dgvUser("Nam", e.RowIndex).Value)

        'ラベルに反映
        selectUserLabel.Text = selectedNam

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

    Private Sub rbtn_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtnHonkan.CheckedChanged, rbtnAnex.CheckedChanged
        Dim rbtn As RadioButton = CType(sender, RadioButton)
        Dim rbtnText As String = rbtn.Text

        If rbtn.Checked = True Then
            rbtn.BackColor = rbtnBackColor
            If rbtnText = "本館" Then
                '選択年月の本館のデータ表示
                '

                '測定日のラベル部分設定
                settingUnitLabel(rbtnText)

            ElseIf rbtnText = "ｱﾈｯｸｽ" Then
                '選択年月のｱﾈｯｸｽのデータ表示
                '

                '測定日のラベル部分設定
                settingUnitLabel(rbtnText)

            End If
        Else
            rbtn.BackColor = Color.FromKnownColor(KnownColor.Control)
        End If
    End Sub
End Class