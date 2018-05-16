Imports System.Data.OleDb

Public Class 体重管理

    '選択している名前保持用
    Private selectedNam As String = ""

    '選択しているセルの行のインデックス保持用
    Private selectedRowIndex As Integer = 0

    '選択しているセルの列のインデックス保持用
    Private selectedColumnIndex As Integer = 0

    'ラジオボタン背景色
    Private rbtnBackColor As Color = Color.FromArgb(255, 255, 0)

    '名前表示色
    Private displayNameColor As Color = Color.FromArgb(0, 120, 215)

    'ユニット左側
    Private unitLeftDt As DataTable
    Private unitLeftAdapter As OleDbDataAdapter

    'ユニット右側
    Private unitRightDt As DataTable
    Private unitRightAdapter As OleDbDataAdapter

    '最大表示行数
    Private Const MAX_ROW_COUNT As Integer = 30

    '室号列のセルスタイル
    Private roomColumnCellStyle As DataGridViewCellStyle

    '体重列のセルスタイル
    Private weitColumnCellStyle As DataGridViewCellStyle

    '前月比のセルスタイル
    Private cmprColumnCellStyle As DataGridViewCellStyle

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

    Private Sub 体重管理_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Left = 0
        Me.Top = 55
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        selectUserLabel.ForeColor = displayNameColor
        initYmdBox()
        createCellStyle()

        rbtnHonkan.Checked = True

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

    Private Sub displayUnitLeftData(ymStr As String, div As Integer)
        dgvUnitLeft.Columns.Clear()

        '表示前設定
        settingDgvUnit(dgvUnitLeft)

        'データ取得、表示
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        unitLeftAdapter = New OleDbDataAdapter(SQLCm)
        unitLeftDt = New DataTable()
        SQLCm.CommandText = "select Ym, Div, Gyo, Nam, Weit, Hana, Mori, Hosi, Sora, Umi, Tuki, Cmpr from Weit where Ym=@ymStr and Div=@div and Gyo <= 30 order by Gyo"
        SQLCm.Parameters.Clear()
        SQLCm.Parameters.Add("@ymStr", OleDbType.Char).Value = ymStr
        SQLCm.Parameters.Add("@div", OleDbType.Integer).Value = div
        unitLeftAdapter.Fill(unitLeftDt)

        'insertコマンド
        unitLeftAdapter.InsertCommand = New OleDbCommand("insert into Weit (Ym, Div, Gyo, Nam, Weit, Hana, Mori, Hosi, Sora, Umi, Tuki, Cmpr) values (@ym, @div, @gyo, @nam, @weit, @hana, @mori, @hosi, @sora, @umi, tuki, @cmpr)", Cn)
        unitLeftAdapter.InsertCommand.Parameters.Add("@ym", OleDbType.Char, 7, "Ym")
        unitLeftAdapter.InsertCommand.Parameters.Add("@div", OleDbType.Integer, 1, "Div")
        unitLeftAdapter.InsertCommand.Parameters.Add("@gyo", OleDbType.Integer, 1, "Gyo")
        unitLeftAdapter.InsertCommand.Parameters.Add("@nam", OleDbType.Char, 15, "Nam")
        unitLeftAdapter.InsertCommand.Parameters.Add("@weit", OleDbType.Decimal, 10, "Weit")
        unitLeftAdapter.InsertCommand.Parameters.Add("@hana", OleDbType.Char, 10, "Hana")
        unitLeftAdapter.InsertCommand.Parameters.Add("@mori", OleDbType.Char, 10, "Mori")
        unitLeftAdapter.InsertCommand.Parameters.Add("@hosi", OleDbType.Char, 10, "Hosi")
        unitLeftAdapter.InsertCommand.Parameters.Add("@sora", OleDbType.Char, 10, "Sora")
        unitLeftAdapter.InsertCommand.Parameters.Add("@umi", OleDbType.Char, 10, "Umi")
        unitLeftAdapter.InsertCommand.Parameters.Add("@tuki", OleDbType.Char, 10, "Tuki")
        unitLeftAdapter.InsertCommand.Parameters.Add("@cmpr", OleDbType.Decimal, 10, "Cmpr")

        'updateコマンド
        unitLeftAdapter.UpdateCommand = New OleDbCommand("update Weit set Nam=@nam, Weit=@weit, Hana=@hana, Mori=@mori, Hosi=@hosi, Sora=@sora, Umi=@umi, Tuki=@tuki, Cmpr=@cmpr where Ym=@ym and Div=@div and Gyo=@gyo")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@nam", OleDbType.Char, 15, "Nam")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@weit", OleDbType.Decimal, 10, "Weit")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@hana", OleDbType.Char, 10, "Hana")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@mori", OleDbType.Char, 10, "Mori")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@hosi", OleDbType.Char, 10, "Hosi")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@sora", OleDbType.Char, 10, "Sora")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@umi", OleDbType.Char, 10, "Umi")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@tuki", OleDbType.Char, 10, "Tuki")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@cmpr", OleDbType.Decimal, 10, "Cmpr")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@ym", OleDbType.Char, 7, "Ym")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@div", OleDbType.Integer, 1, "Div")
        unitLeftAdapter.UpdateCommand.Parameters.Add("@gyo", OleDbType.Integer, 1, "Gyo")

        addBlankRow(unitLeftDt)
        addRoomColumn(unitLeftDt)
        addRoomNum4UnitLeft()

        dgvUnitLeft.DataSource = unitLeftDt

        '表示後設定
        settingDgvUnitColumns(dgvUnitLeft)

        '測定日部分設定
        settingYmdBox(unitLeftDt)

    End Sub

    Private Sub displayUnitRightData(ymStr As String, div As Integer)
        dgvUnitRight.Columns.Clear()

        '表示前設定
        settingDgvUnit(dgvUnitRight)

        'データ取得、表示
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        unitRightAdapter = New OleDbDataAdapter(SQLCm)
        unitRightDt = New DataTable()
        SQLCm.CommandText = "select Ym, Div, Gyo, Nam, Weit, Hana, Mori, Hosi, Sora, Umi, Tuki, Cmpr from Weit where Ym=@ymStr and Div=@div and Gyo >= 31 order by Gyo"
        SQLCm.Parameters.Clear()
        SQLCm.Parameters.Add("@ymStr", OleDbType.Char).Value = ymStr
        SQLCm.Parameters.Add("@div", OleDbType.Integer).Value = div
        unitRightAdapter.Fill(unitRightDt)

        'insertコマンド
        unitRightAdapter.InsertCommand = New OleDbCommand("insert into Weit (Ym, Div, Gyo, Nam, Weit, Hana, Mori, Hosi, Sora, Umi, Tuki, Cmpr) values (@ym, @div, @gyo, @nam, @weit, @hana, @mori, @hosi, @sora, @umi, tuki, @cmpr)", Cn)
        unitRightAdapter.InsertCommand.Parameters.Add("@ym", OleDbType.Char, 7, "Ym")
        unitRightAdapter.InsertCommand.Parameters.Add("@div", OleDbType.Integer, 1, "Div")
        unitRightAdapter.InsertCommand.Parameters.Add("@gyo", OleDbType.Integer, 1, "Gyo")
        unitRightAdapter.InsertCommand.Parameters.Add("@nam", OleDbType.Char, 15, "Nam")
        unitRightAdapter.InsertCommand.Parameters.Add("@weit", OleDbType.Decimal, 10, "Weit")
        unitRightAdapter.InsertCommand.Parameters.Add("@hana", OleDbType.Char, 10, "Hana")
        unitRightAdapter.InsertCommand.Parameters.Add("@mori", OleDbType.Char, 10, "Mori")
        unitRightAdapter.InsertCommand.Parameters.Add("@hosi", OleDbType.Char, 10, "Hosi")
        unitRightAdapter.InsertCommand.Parameters.Add("@sora", OleDbType.Char, 10, "Sora")
        unitRightAdapter.InsertCommand.Parameters.Add("@umi", OleDbType.Char, 10, "Umi")
        unitRightAdapter.InsertCommand.Parameters.Add("@tuki", OleDbType.Char, 10, "Tuki")
        unitRightAdapter.InsertCommand.Parameters.Add("@cmpr", OleDbType.Decimal, 10, "Cmpr")

        'updateコマンド
        unitRightAdapter.UpdateCommand = New OleDbCommand("update Weit set Nam=@nam, Weit=@weit, Hana=@hana, Mori=@mori, Hosi=@hosi, Sora=@sora, Umi=@umi, Tuki=@tuki, Cmpr=@cmpr where Ym=@ym and Div=@div and Gyo=@gyo")
        unitRightAdapter.UpdateCommand.Parameters.Add("@nam", OleDbType.Char, 15, "Nam")
        unitRightAdapter.UpdateCommand.Parameters.Add("@weit", OleDbType.Decimal, 10, "Weit")
        unitRightAdapter.UpdateCommand.Parameters.Add("@hana", OleDbType.Char, 10, "Hana")
        unitRightAdapter.UpdateCommand.Parameters.Add("@mori", OleDbType.Char, 10, "Mori")
        unitRightAdapter.UpdateCommand.Parameters.Add("@hosi", OleDbType.Char, 10, "Hosi")
        unitRightAdapter.UpdateCommand.Parameters.Add("@sora", OleDbType.Char, 10, "Sora")
        unitRightAdapter.UpdateCommand.Parameters.Add("@umi", OleDbType.Char, 10, "Umi")
        unitRightAdapter.UpdateCommand.Parameters.Add("@tuki", OleDbType.Char, 10, "Tuki")
        unitRightAdapter.UpdateCommand.Parameters.Add("@cmpr", OleDbType.Decimal, 10, "Cmpr")
        unitRightAdapter.UpdateCommand.Parameters.Add("@ym", OleDbType.Char, 7, "Ym")
        unitRightAdapter.UpdateCommand.Parameters.Add("@div", OleDbType.Integer, 1, "Div")
        unitRightAdapter.UpdateCommand.Parameters.Add("@gyo", OleDbType.Integer, 1, "Gyo")

        addBlankRow(unitRightDt)
        addRoomColumn(unitRightDt)
        addRoomNum4UnitRight()

        dgvUnitRight.DataSource = unitRightDt

        '表示後設定
        settingDgvUnitColumns(dgvUnitRight)
    End Sub

    Private Sub addRoomColumn(dt As DataTable)
        Dim roomColumn As New DataColumn()
        roomColumn.ColumnName = "Room"
        roomColumn.DataType = System.Type.GetType("System.Int32")
        dt.Columns.Add(roomColumn)
    End Sub

    Private Sub addRoomNum4UnitLeft()
        Dim numArray() As Integer = {1, 2, 3, 5, 6, 7, 8, 10, 11, 12}
        For i As Integer = 0 To MAX_ROW_COUNT - 1
            unitLeftDt.Rows(i)("Room") = 100 * (i \ 10 + 1) + numArray(i Mod 10)
        Next
    End Sub

    Private Sub addRoomNum4UnitRight()
        Dim numArray() As Integer = {1, 2, 3, 5, 6, 7, 8, 10, 11, 12}
        For i As Integer = 0 To MAX_ROW_COUNT - 1
            unitRightDt.Rows(i)("Room") = 100 * (i \ 10 + 1) + numArray(i Mod 10) + 400
        Next
    End Sub

    Private Sub addBlankRow(dt As DataTable)
        Dim rowCount As Integer = dt.Rows.Count
        If rowCount < MAX_ROW_COUNT Then
            Dim row As DataRow
            For i As Integer = rowCount To MAX_ROW_COUNT - 1
                row = dt.NewRow()
                dt.Rows.Add(row)
            Next
        End If
    End Sub

    Private Sub createCellStyle()
        '室号列のセルスタイル
        roomColumnCellStyle = New DataGridViewCellStyle()
        roomColumnCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
        roomColumnCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        '体重列のセルスタイル
        weitColumnCellStyle = New DataGridViewCellStyle()
        weitColumnCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        weitColumnCellStyle.Format = "#,##0.00"

        '前月比のセルスタイル
        cmprColumnCellStyle = New DataGridViewCellStyle()
        cmprColumnCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        cmprColumnCellStyle.Format = "#,##0.00"

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
            .ScrollBars = ScrollBars.Vertical
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

    Private Sub settingDgvUnit(dgv As DataGridView)
        'DoubleBufferedプロパティをTrue
        TopForm.EnableDoubleBuffering(dgv)

        'dgv設定
        With dgv
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .MultiSelect = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .RowHeadersVisible = False
            .RowTemplate.Height = 16
            .ColumnHeadersHeight = 20
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .EnableHeadersVisualStyles = False
            .ColumnHeadersDefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ShowCellToolTips = False
            .ScrollBars = ScrollBars.None
        End With
    End Sub

    Private Sub settingDgvUnitColumns(dgv As DataGridView)
        '並び替えができないようにする
        For Each c As DataGridViewColumn In dgv.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        With dgv
            '非表示
            .Columns("Ym").Visible = False
            .Columns("Div").Visible = False
            .Columns("Gyo").Visible = False
            .Columns("Hana").Visible = False
            .Columns("Mori").Visible = False
            .Columns("Hosi").Visible = False
            .Columns("Sora").Visible = False
            .Columns("Umi").Visible = False
            .Columns("Tuki").Visible = False

            '室号
            With .Columns("Room")
                .Width = 50
                .DefaultCellStyle = roomColumnCellStyle
                .ReadOnly = True
                .DisplayIndex = 0
                .HeaderText = "室号"
            End With

            '氏名
            With .Columns("Nam")
                .Width = 100
                .ReadOnly = True
                .HeaderText = "氏 名"
            End With

            '体重
            With .Columns("Weit")
                .Width = 60
                .HeaderText = "体 重"
                .DefaultCellStyle = weitColumnCellStyle
            End With

            '前月比
            With .Columns("Cmpr")
                .Width = 60
                .ReadOnly = True
                .HeaderText = "前月比"
                .DefaultCellStyle = cmprColumnCellStyle
            End With
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

    Private Sub settingYmdBox(dt As DataTable)
        Dim row As DataRow = If(IsNothing(dt.Rows(0)), Nothing, dt.Rows(0))
        If IsNothing(row) Then
            Return
        End If

        If row("Sora") <> "" Then
            ymdBoxUnit1.setADStr(row("Sora"))
        End If
        If row("Mori") <> "" Then
            ymdBoxUnit2.setADStr(row("Mori"))
        End If
        If row("Hosi") <> "" Then
            ymdBoxUnit3.setADStr(row("Hosi"))
        End If
        If row("Tuki") <> "" Then
            ymdBoxUnit4.setADStr(row("Tuki"))
        End If
        If row("Hana") <> "" Then
            ymdBoxUnit5.setADStr(row("Hana"))
        End If
        If row("Umi") <> "" Then
            ymdBoxUnit6.setADStr(row("Umi"))
        End If
    End Sub

    Private Sub settingAddedRowState(dt As DataTable)
        For Each row As DataRow In dt.Rows
            row.SetAdded()
        Next
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

    Private Sub dgvUnit_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvUnitLeft.CellMouseClick, dgvUnitRight.CellMouseClick
        Dim dgv As DataGridView = CType(sender, DataGridView)
        selectedRowIndex = dgv.CurrentCell.RowIndex
        selectedColumnIndex = dgv.CurrentCell.ColumnIndex
        If selectedNam <> "" AndAlso dgv.Columns(e.ColumnIndex).Name = "Nam" Then
            dgv(e.ColumnIndex, e.RowIndex).Value = selectedNam
        End If
    End Sub

    Private Sub dgvUnit_CellMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvUnitLeft.CellMouseDoubleClick, dgvUnitRight.CellMouseDoubleClick
        Dim dgv As DataGridView = CType(sender, DataGridView)
        If e.RowIndex > -1 AndAlso dgv.Columns(e.ColumnIndex).Name = "Room" Then
            dgv("Nam", e.RowIndex).Value = ""
            dgv("Weit", e.RowIndex).Value = DBNull.Value
            dgv("Cmpr", e.RowIndex).Value = DBNull.Value
        End If
    End Sub

    Private Sub dgvUnit_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvUnitLeft.CellPainting, dgvUnitRight.CellPainting
        If (e.RowIndex Mod 10 = 0) AndAlso (e.PaintParts And DataGridViewPaintParts.Border) = DataGridViewPaintParts.Border Then
            e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.Inset
        End If
    End Sub

    Private Sub rbtn_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtnHonkan.CheckedChanged, rbtnAnex.CheckedChanged
        '左上の名前ラベルのクリア
        selectUserLabel.Text = ""

        Dim rbtn As RadioButton = CType(sender, RadioButton)
        Dim rbtnText As String = rbtn.Text

        If rbtn.Checked = True Then
            rbtn.BackColor = rbtnBackColor
            If rbtnText = "本館" Then
                '選択年月の本館のデータ表示
                displayUnitLeftData("2018/04", 1)
                displayUnitRightData("2018/04", 1)

                '測定日のラベル部分設定
                settingUnitLabel(rbtnText)

            ElseIf rbtnText = "ｱﾈｯｸｽ" Then
                '選択年月のｱﾈｯｸｽのデータ表示
                displayUnitLeftData("2018/04", 2)
                displayUnitRightData("2018/04", 2)

                '測定日のラベル部分設定
                settingUnitLabel(rbtnText)

            End If
        Else
            rbtn.BackColor = Color.FromKnownColor(KnownColor.Control)
        End If
    End Sub

    Private Sub deleteMonthData(ymStr As String, div As Integer)
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        SQLCm.CommandText = "delete from Weit where Ym=@ym and Div=@div"
        SQLCm.Parameters.Clear()
        SQLCm.Parameters.Add("@ym", OleDbType.Char).Value = ymStr
        SQLCm.Parameters.Add("@div", OleDbType.Integer).Value = div
        cn.Open()
        SQLCm.ExecuteNonQuery()
        cn.Close()
    End Sub

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        '全ての行の状態をAddedに
        unitLeftDt.AcceptChanges()
        unitRightDt.AcceptChanges()
        settingAddedRowState(unitLeftDt)
        settingAddedRowState(unitRightDt)

        'delete
        Dim div As Integer = If(rbtnHonkan.Checked, 1, 2)
        deleteMonthData("2018/04", div)

        'insert
        unitLeftAdapter.Update(unitLeftDt)
        unitRightAdapter.Update(unitRightDt)

        '再表示
        displayUnitLeftData("2018/04", div)
        displayUnitRightData("2018/04", div)

    End Sub

    Private Sub dgvUnitLeft_GotFocus(sender As Object, e As System.EventArgs) Handles dgvUnitLeft.GotFocus
        dgvUnitRight.ClearSelection()
    End Sub

    Private Sub dgvUnitRight_GotFocus(sender As Object, e As System.EventArgs) Handles dgvUnitRight.GotFocus
        dgvUnitLeft.ClearSelection()
    End Sub

    Private Sub dgvUnitLeft_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles dgvUnitLeft.KeyDown
        If e.KeyCode = Keys.Right AndAlso dgvUnitLeft.Columns(selectedColumnIndex).Name = "Cmpr" Then
            dgvUnitLeft.ClearSelection()
            dgvUnitRight.Focus()
            dgvUnitRight(12, selectedRowIndex).Selected = True
        End If
    End Sub
End Class