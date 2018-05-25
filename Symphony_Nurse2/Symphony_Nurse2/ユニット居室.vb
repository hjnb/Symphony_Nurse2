Imports System.Data.OleDb

Public Class ユニット居室

    '選択している名前（漢字）保持用
    Private selectedNam As String = ""

    '選択している名前（カナ）保持用
    Private selectedKana As String = ""

    '選択しているID保持用
    Private selectedID As String = ""

    Private unitDt As DataTable
    Private unitAdapter As OleDbDataAdapter

    '"○の家"の行のセルスタイル
    Private houseNameCellStyle As DataGridViewCellStyle

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

        'セルスタイル作成
        createCellStyle()

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
        dgvUnit.Columns.Clear()

        '表示前設定
        settingDgvUnit()

        'データ取得、表示
        Dim moriList As New List(Of String)
        Dim hosiList As New List(Of String)
        Dim tukiList As New List(Of String)
        Dim hanaList As New List(Of String)
        Dim okaList As New List(Of String)
        Dim nijiList As New List(Of String)
        Dim hikariList As New List(Of String)
        Dim yukiList As New List(Of String)
        Dim kazeList As New List(Of String)
        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        unitAdapter = New OleDbDataAdapter(SQLCm)
        unitDt = New DataTable()
        SQLCm.CommandText = "select * from Unit order by Gyo"
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        While reader.Read() = True
            moriList.Add(reader("Mori"))
            hosiList.Add(reader("Hosi"))
            tukiList.Add(reader("Tuki"))
            hanaList.Add(reader("Hana"))
            okaList.Add(reader("Oka"))
            nijiList.Add(reader("Niji"))
            hikariList.Add(reader("Hikari"))
            yukiList.Add(reader("Yuki"))
            kazeList.Add(reader("Kaze"))
        End While
        reader.Close()
        Cn.Close()
        unitAdapter.Fill(unitDt)

        '更新コマンド
        unitAdapter.UpdateCommand = New OleDbCommand("update Unit set SoraID=@soraId, Sora=@sora, SKana=@sKana, MoriID=@moriId, Mori=@mori, MKana=@mKana, HosiID=@hosiId, Hosi=@hosi, HoKana=@hoKana, TukiID=@tukiId, Tuki=@tuki, TKana=@tKana, HanaID=@hanaId, Hana=@hana, HaKana=@haKana, OkaID=@okaId, Oka=@oka, OKana=@oKana, NijiID=@nijiId, Niji=@niji, NiKana=@niKana, HikariID=@hikariId, Hikari=@hikari, HiKana=@hiKana, YukiID=@yukiId, Yuki=@yuki, YuKana=@yuKana, KazeID=@kazeId, Kaze=@kaze, KaKana=@kaKana where Gyo=@gyo", Cn)
        unitAdapter.UpdateCommand.Parameters.Add("@soraId", OleDbType.VarChar, 10, "SoraID")
        unitAdapter.UpdateCommand.Parameters.Add("@sora", OleDbType.VarChar, 10, "Sora")
        unitAdapter.UpdateCommand.Parameters.Add("@sKana", OleDbType.VarChar, 10, "SKana")
        unitAdapter.UpdateCommand.Parameters.Add("@moriId", OleDbType.VarChar, 10, "MoriID")
        unitAdapter.UpdateCommand.Parameters.Add("@mori", OleDbType.VarChar, 10, "Mori")
        unitAdapter.UpdateCommand.Parameters.Add("@mKana", OleDbType.VarChar, 10, "MKana")
        unitAdapter.UpdateCommand.Parameters.Add("@hosiId", OleDbType.VarChar, 10, "HosiID")
        unitAdapter.UpdateCommand.Parameters.Add("@hosi", OleDbType.VarChar, 10, "Hosi")
        unitAdapter.UpdateCommand.Parameters.Add("@hoKana", OleDbType.VarChar, 10, "HoKana")
        unitAdapter.UpdateCommand.Parameters.Add("@tukiId", OleDbType.VarChar, 10, "TukiID")
        unitAdapter.UpdateCommand.Parameters.Add("@tuki", OleDbType.VarChar, 10, "Tuki")
        unitAdapter.UpdateCommand.Parameters.Add("@tKana", OleDbType.VarChar, 10, "TKana")
        unitAdapter.UpdateCommand.Parameters.Add("@hanaId", OleDbType.VarChar, 10, "HanaID")
        unitAdapter.UpdateCommand.Parameters.Add("@hana", OleDbType.VarChar, 10, "Hana")
        unitAdapter.UpdateCommand.Parameters.Add("@haKana", OleDbType.VarChar, 10, "HaKana")
        unitAdapter.UpdateCommand.Parameters.Add("@okaId", OleDbType.VarChar, 10, "OkaID")
        unitAdapter.UpdateCommand.Parameters.Add("@oka", OleDbType.VarChar, 10, "Oka")
        unitAdapter.UpdateCommand.Parameters.Add("@oKana", OleDbType.VarChar, 10, "OKana")
        unitAdapter.UpdateCommand.Parameters.Add("@nijiId", OleDbType.VarChar, 10, "NijiID")
        unitAdapter.UpdateCommand.Parameters.Add("@niji", OleDbType.VarChar, 10, "Niji")
        unitAdapter.UpdateCommand.Parameters.Add("@niKana", OleDbType.VarChar, 10, "NiKana")
        unitAdapter.UpdateCommand.Parameters.Add("@hikariId", OleDbType.VarChar, 10, "HikariID")
        unitAdapter.UpdateCommand.Parameters.Add("@hikari", OleDbType.VarChar, 10, "Hikari")
        unitAdapter.UpdateCommand.Parameters.Add("@hiKana", OleDbType.VarChar, 10, "HiKana")
        unitAdapter.UpdateCommand.Parameters.Add("@yukiId", OleDbType.VarChar, 10, "YukiID")
        unitAdapter.UpdateCommand.Parameters.Add("@yuki", OleDbType.VarChar, 10, "Yuki")
        unitAdapter.UpdateCommand.Parameters.Add("@yuKana", OleDbType.VarChar, 10, "YuKana")
        unitAdapter.UpdateCommand.Parameters.Add("@kazeId", OleDbType.VarChar, 10, "KazeID")
        unitAdapter.UpdateCommand.Parameters.Add("@kaze", OleDbType.VarChar, 10, "Kaze")
        unitAdapter.UpdateCommand.Parameters.Add("@kaKana", OleDbType.VarChar, 10, "KaKana")
        unitAdapter.UpdateCommand.Parameters.Add("@gyo", OleDbType.Integer, 5, "Gyo").SourceVersion = DataRowVersion.Original

        '部屋番号の列追加
        addRoomNumColumn()

        '森以降のユニットのリストをdatatableに追加
        addUnitNamListRow(moriList, hosiList, tukiList, hanaList, okaList, nijiList, hikariList, yukiList, kazeList)

        'dgvに設定、表示
        dgvUnit.DataSource = unitDt

        '表示後設定
        settingDgvUnitColumns()
    End Sub

    Private Sub addRoomNumColumn()
        Dim roomNumColumn As DataColumn = New DataColumn()
        roomNumColumn.ColumnName = "RoomNum"
        roomNumColumn.DataType = System.Type.GetType("System.Int32")
        unitDt.Columns.Add(roomNumColumn)
    End Sub

    Private Sub addUnitNamListRow(moriList As List(Of String), hosiList As List(Of String), tukiList As List(Of String), hanaList As List(Of String), okaList As List(Of String), nijiList As List(Of String), hikariList As List(Of String), yukiList As List(Of String), kazeList As List(Of String))
        Dim row As DataRow
        Dim roomNumColumnName As String = "RoomNum"
        Dim namInputColumnName As String = "Sora"
        Dim roomNum() As Integer = {1, 2, 3, 5, 6, 7, 8, 10, 11, 12}
        Const SORA_ROOM As Integer = 100
        Const MORI_ROOM As Integer = 200
        Const HOSI_ROOM As Integer = 300
        Const TUKI_ROOM As Integer = 500
        Const HANA_ROOM As Integer = 600
        Const OKA_ROOM As Integer = 100
        Const NIJI_ROOM As Integer = 200
        Const HIKARI_ROOM As Integer = 300
        Const YUKI_ROOM As Integer = 500
        Const KAZE_ROOM As Integer = 600

        '空の家
        row = unitDt.NewRow()
        row(namInputColumnName) = "空の家"
        unitDt.Rows.InsertAt(row, 0)
        For i As Integer = 1 To 10
            unitDt.Rows(i)(roomNumColumnName) = SORA_ROOM + roomNum(i - 1)
        Next

        '森の家
        row = unitDt.NewRow()
        row(namInputColumnName) = "森の家"
        unitDt.Rows.Add(row)
        For i As Integer = 0 To 9
            row = unitDt.NewRow()
            row(roomNumColumnName) = MORI_ROOM + roomNum(i)
            row(namInputColumnName) = moriList.Item(i)
            unitDt.Rows.Add(row)
        Next

        '星の家
        row = unitDt.NewRow()
        row(namInputColumnName) = "星の家"
        unitDt.Rows.Add(row)
        For i As Integer = 0 To 9
            row = unitDt.NewRow()
            row(roomNumColumnName) = HOSI_ROOM + roomNum(i)
            row(namInputColumnName) = hosiList.Item(i)
            unitDt.Rows.Add(row)
        Next

        '月の家
        row = unitDt.NewRow()
        row(namInputColumnName) = "月の家"
        unitDt.Rows.Add(row)
        For i As Integer = 0 To 9
            row = unitDt.NewRow()
            row(roomNumColumnName) = TUKI_ROOM + roomNum(i)
            row(namInputColumnName) = tukiList.Item(i)
            unitDt.Rows.Add(row)
        Next

        '花の家
        row = unitDt.NewRow()
        row(namInputColumnName) = "花の家"
        unitDt.Rows.Add(row)
        For i As Integer = 0 To 9
            row = unitDt.NewRow()
            row(roomNumColumnName) = HANA_ROOM + roomNum(i)
            row(namInputColumnName) = hanaList.Item(i)
            unitDt.Rows.Add(row)
        Next

        '丘の家
        row = unitDt.NewRow()
        row(namInputColumnName) = "丘の家"
        unitDt.Rows.Add(row)
        For i As Integer = 0 To 9
            row = unitDt.NewRow()
            row(roomNumColumnName) = OKA_ROOM + roomNum(i)
            row(namInputColumnName) = okaList.Item(i)
            unitDt.Rows.Add(row)
        Next

        '虹の家
        row = unitDt.NewRow()
        row(namInputColumnName) = "虹の家"
        unitDt.Rows.Add(row)
        For i As Integer = 0 To 9
            row = unitDt.NewRow()
            row(roomNumColumnName) = NIJI_ROOM + roomNum(i)
            row(namInputColumnName) = nijiList.Item(i)
            unitDt.Rows.Add(row)
        Next

        '光の家
        row = unitDt.NewRow()
        row(namInputColumnName) = "光の家"
        unitDt.Rows.Add(row)
        For i As Integer = 0 To 9
            row = unitDt.NewRow()
            row(roomNumColumnName) = HIKARI_ROOM + roomNum(i)
            row(namInputColumnName) = hikariList.Item(i)
            unitDt.Rows.Add(row)
        Next

        '雪の家
        row = unitDt.NewRow()
        row(namInputColumnName) = "雪の家"
        unitDt.Rows.Add(row)
        For i As Integer = 0 To 9
            row = unitDt.NewRow()
            row(roomNumColumnName) = YUKI_ROOM + roomNum(i)
            row(namInputColumnName) = yukiList.Item(i)
            unitDt.Rows.Add(row)
        Next

        '風の家
        row = unitDt.NewRow()
        row(namInputColumnName) = "風の家"
        unitDt.Rows.Add(row)
        For i As Integer = 0 To 9
            row = unitDt.NewRow()
            row(roomNumColumnName) = KAZE_ROOM + roomNum(i)
            row(namInputColumnName) = kazeList.Item(i)
            unitDt.Rows.Add(row)
        Next

    End Sub

    Private Sub createCellStyle()
        houseNameCellStyle = New DataGridViewCellStyle()
        With houseNameCellStyle
            .BackColor = Color.FromKnownColor(KnownColor.Control) '背景色
            .SelectionBackColor = Color.FromKnownColor(KnownColor.Control) '選択時の背景色
            .SelectionForeColor = Color.Black '選択時の前景色
        End With
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

        'dgv設定
        With dgvUnit
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .MultiSelect = False
            .RowHeadersVisible = False
            .ColumnHeadersVisible = False
            .RowTemplate.Height = 15
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .ShowCellToolTips = False
        End With
    End Sub

    Private Sub settingDgvUnitColumns()
        With dgvUnit

            '表示
            '.Columns("Sora").Visible = True
            '.Columns("RoomNum").Visible = True

            '非表示
            .Columns("Gyo").Visible = False
            .Columns("SoraID").Visible = False
            .Columns("SKana").Visible = False
            .Columns("MoriID").Visible = False
            .Columns("Mori").Visible = False
            .Columns("MKana").Visible = False
            .Columns("HosiID").Visible = False
            .Columns("Hosi").Visible = False
            .Columns("HoKana").Visible = False
            .Columns("TukiID").Visible = False
            .Columns("Tuki").Visible = False
            .Columns("TKana").Visible = False
            .Columns("HanaID").Visible = False
            .Columns("Hana").Visible = False
            .Columns("HaKana").Visible = False
            .Columns("OkaID").Visible = False
            .Columns("Oka").Visible = False
            .Columns("OKana").Visible = False
            .Columns("NijiID").Visible = False
            .Columns("Niji").Visible = False
            .Columns("NiKana").Visible = False
            .Columns("HikariID").Visible = False
            .Columns("Hikari").Visible = False
            .Columns("HiKana").Visible = False
            .Columns("YukiID").Visible = False
            .Columns("Yuki").Visible = False
            .Columns("YuKana").Visible = False
            .Columns("KazeID").Visible = False
            .Columns("Kaze").Visible = False
            .Columns("KaKana").Visible = False

            With .Columns("RoomNum")
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DisplayIndex = 0
                .ReadOnly = True
                .Width = 30
            End With

            With .Columns("Sora")
                .Width = 85
            End With

            '"○の家"の行のセルスタイルを設定
            For i As Integer = 0 To 99 Step 11
                dgvUnit.Rows(i).DefaultCellStyle = houseNameCellStyle
                dgvUnit.Rows(i).ReadOnly = True
            Next

        End With
    End Sub

    Private Sub setModifiedState()
        For i As Integer = 1 To 10
            unitDt.Rows(i).SetModified()
        Next
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
        selectedID = checkDBNullValue(dgvUser("Id", e.RowIndex).Value)
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

    Private Sub dgvUnit_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvUnit.CellMouseClick
        If e.RowIndex Mod 11 <> 0 AndAlso selectedNam <> "" Then
            dgvUnit("Sora", e.RowIndex).Value = selectedNam
            dgvUnit("SKana", e.RowIndex).Value = selectedKana
            dgvUnit("SoraID", e.RowIndex).Value = selectedID
        End If

        If e.RowIndex Mod 11 <> 0 AndAlso dgvUnit.Columns(e.ColumnIndex).Name = "Sora" Then
            dgvUnit.BeginEdit(True)
        End If
    End Sub

    Private Sub dgvUnit_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvUnit.CellPainting
        If e.RowIndex = 55 AndAlso (e.PaintParts And DataGridViewPaintParts.Border) = DataGridViewPaintParts.Border Then
            e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.Inset
        End If
    End Sub

    Private Sub dgvUnit_CellValueChanged(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvUnit.CellValueChanged
        If e.RowIndex <= 11 Then
            '空の家
            If dgvUnit.Columns(e.ColumnIndex).Name = "Sora" Then
                Dim inputStr As String = checkDBNullValue(dgvUnit(e.ColumnIndex, e.RowIndex).Value)
                If inputStr = "" Then
                    dgvUnit("SoraID", e.RowIndex).Value = ""
                    dgvUnit("SKana", e.RowIndex).Value = ""
                End If
            End If
        ElseIf e.RowIndex > 11 Then
            '森の家以降
            Dim q As Integer = e.RowIndex \ 11 '商
            Dim r As Integer = e.RowIndex Mod 11 '余り
            Dim changedColumnName As String = dgvUnit.Columns(e.ColumnIndex).Name
            Dim inputIdColumnName As String = "SoraID"
            Dim inputNamColumnName As String = "Sora"
            Dim inputKanaColumnName As String = "SKana"
            Dim inputStr As String = checkDBNullValue(dgvUnit(e.ColumnIndex, e.RowIndex).Value)

            If q = 1 Then
                '森の家
                If inputStr = "" Then
                    dgvUnit("MoriID", r).Value = ""
                    dgvUnit("Mori", r).Value = ""
                    dgvUnit("MKana", r).Value = ""
                Else
                    If changedColumnName = inputIdColumnName Then
                        dgvUnit("MoriID", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputNamColumnName Then
                        dgvUnit("Mori", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputKanaColumnName Then
                        dgvUnit("MKana", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            ElseIf q = 2 Then
                '星の家
                If inputStr = "" Then
                    dgvUnit("HosiID", r).Value = ""
                    dgvUnit("Hosi", r).Value = ""
                    dgvUnit("HoKana", r).Value = ""
                Else
                    If changedColumnName = inputIdColumnName Then
                        dgvUnit("HosiID", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputNamColumnName Then
                        dgvUnit("Hosi", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputKanaColumnName Then
                        dgvUnit("HoKana", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            ElseIf q = 3 Then
                '月の家
                If inputStr = "" Then
                    dgvUnit("TukiID", r).Value = ""
                    dgvUnit("Tuki", r).Value = ""
                    dgvUnit("TKana", r).Value = ""
                Else
                    If changedColumnName = inputIdColumnName Then
                        dgvUnit("TukiID", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputNamColumnName Then
                        dgvUnit("Tuki", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputKanaColumnName Then
                        dgvUnit("TKana", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            ElseIf q = 4 Then
                '花の家
                If inputStr = "" Then
                    dgvUnit("HanaID", r).Value = ""
                    dgvUnit("Hana", r).Value = ""
                    dgvUnit("HaKana", r).Value = ""
                Else
                    If changedColumnName = inputIdColumnName Then
                        dgvUnit("HanaID", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputNamColumnName Then
                        dgvUnit("Hana", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputKanaColumnName Then
                        dgvUnit("HaKana", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            ElseIf q = 5 Then
                '丘の家
                If inputStr = "" Then
                    dgvUnit("OkaID", r).Value = ""
                    dgvUnit("Oka", r).Value = ""
                    dgvUnit("OKana", r).Value = ""
                Else
                    If changedColumnName = inputIdColumnName Then
                        dgvUnit("OkaID", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputNamColumnName Then
                        dgvUnit("Oka", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputKanaColumnName Then
                        dgvUnit("OKana", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            ElseIf q = 6 Then
                '虹の家
                If inputStr = "" Then
                    dgvUnit("NijiID", r).Value = ""
                    dgvUnit("Niji", r).Value = ""
                    dgvUnit("NiKana", r).Value = ""
                Else
                    If changedColumnName = inputIdColumnName Then
                        dgvUnit("NijiID", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputNamColumnName Then
                        dgvUnit("Niji", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputKanaColumnName Then
                        dgvUnit("NiKana", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            ElseIf q = 7 Then
                '光の家
                If inputStr = "" Then
                    dgvUnit("HikariID", r).Value = ""
                    dgvUnit("Hikari", r).Value = ""
                    dgvUnit("HiKana", r).Value = ""
                Else
                    If changedColumnName = inputIdColumnName Then
                        dgvUnit("HikariID", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputNamColumnName Then
                        dgvUnit("Hikari", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputKanaColumnName Then
                        dgvUnit("HiKana", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            ElseIf q = 8 Then
                '雪の家
                If inputStr = "" Then
                    dgvUnit("YukiID", r).Value = ""
                    dgvUnit("Yuki", r).Value = ""
                    dgvUnit("YuKana", r).Value = ""
                Else
                    If changedColumnName = inputIdColumnName Then
                        dgvUnit("YukiID", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputNamColumnName Then
                        dgvUnit("Yuki", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputKanaColumnName Then
                        dgvUnit("YuKana", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            ElseIf q = 9 Then
                '風の家
                If inputStr = "" Then
                    dgvUnit("KazeID", r).Value = ""
                    dgvUnit("Kaze", r).Value = ""
                    dgvUnit("KaKana", r).Value = ""
                Else
                    If changedColumnName = inputIdColumnName Then
                        dgvUnit("KazeID", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputNamColumnName Then
                        dgvUnit("Kaze", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    ElseIf changedColumnName = inputKanaColumnName Then
                        dgvUnit("KaKana", r).Value = dgvUnit(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        '一度全ての状態を確定
        unitDt.AcceptChanges()

        '空の家部分の１０行をModified状態に設定
        setModifiedState()

        '更新
        unitAdapter.Update(unitDt.Select(Nothing, Nothing, DataViewRowState.ModifiedCurrent))
        displayUnitMasterData()
        MsgBox("登録しました。")
    End Sub    
End Class