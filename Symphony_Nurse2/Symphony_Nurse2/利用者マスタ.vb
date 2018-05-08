Imports System.Data.OleDb

Public Class 利用者マスタ

    'スクロール位置保持用
    Private scrollPosition As Integer = 0

    '選択行位置保持用
    Private selectedRowIndex As Integer = 0

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

    Private Sub 利用者マスタ_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Left = 0
        Me.Top = 55
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        '入力テキストボックスの設定
        settingInputTextBox()

        'データの表示
        displayUserMasterData(dgvUserMaster)
    End Sub

    Private Sub 利用者マスタ_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If e.Control = False Then
                Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
            End If
        End If
    End Sub

    Private Sub displayUserMasterData(dgv As DataGridView)
        dgv.Columns.Clear()
        settingDgv(dgvUserMaster)

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        SQLCm.CommandText = "select Id, Nam, Kana, Sex, Format(Birth, 'gee/mm/dd') as Birth, Int((Format(NOW(),'YYYYMMDD')-Format(Birth, 'YYYYMMDD'))/10000) as Age, Kaigo, Dsp from KihonM order by Kana"
        Cn.Open()
        Adapter.Fill(Table)
        dgv.DataSource = Table
        Cn.Close()
        Cn.Dispose()

        settingDgvColumns(dgvUserMaster)

        'スクロールの位置設定
        dgvUserMaster.FirstDisplayedScrollingRowIndex = scrollPosition

        '選択行設定
        dgvUserMaster.Rows(selectedRowIndex).Selected = True

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
            .RowTemplate.Height = 18
            .RowHeadersWidth = 35
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .RowTemplate.HeaderCell = New dgvRowHeaderCell() '行ヘッダの三角マークを非表示に
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With

    End Sub

    Private Sub settingDgvColumns(dgv As DataGridView)
        With dgv
            With .Columns("Id")
                .HeaderText = "利用者ID"
                .Width = 60
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            With .Columns("Nam")
                .HeaderText = "漢字氏名"
                .Width = 110
            End With

            With .Columns("Kana")
                .HeaderText = "カナ氏名"
                .Width = 110
            End With

            With .Columns("Sex")
                .HeaderText = "性別"
                .Width = 40
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            With .Columns("Birth")
                .HeaderText = "生年月日"
                .Width = 100
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            With .Columns("Age")
                .HeaderText = "年齢"
                .Width = 40
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            With .Columns("Kaigo")
                .HeaderText = "要介護"
                .Width = 50
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            With .Columns("Dsp")
                .HeaderText = "表示"
                .Width = 40
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

        End With
    End Sub

    Private Sub settingInputTextBox()
        '利用者ID
        idBox.ImeMode = Windows.Forms.ImeMode.Disable

        '漢字氏名
        namBox.ImeMode = Windows.Forms.ImeMode.Hiragana

        'カナ氏名
        kanaBox.ImeMode = Windows.Forms.ImeMode.KatakanaHalf

        '性別
        sexBox.ImeMode = Windows.Forms.ImeMode.Disable
        sexBox.TextAlign = HorizontalAlignment.Center

        '要介護
        kaigoBox.ImeMode = Windows.Forms.ImeMode.Disable
        kaigoBox.TextAlign = HorizontalAlignment.Center

    End Sub

    Private Sub inputClear()
        idBox.Text = ""
        namBox.Text = ""
        kanaBox.Text = ""
        sexBox.Text = ""
        birthYmdBox.clearText()
        kaigoBox.Text = ""
    End Sub

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        Dim nam As String = namBox.Text
        Dim kana As String = StrConv(kanaBox.Text, VbStrConv.Narrow) '半角へ変換
        Dim sex As String = sexBox.Text
        Dim birth As String = birthYmdBox.getADStr()
        Dim kaigo As String = kaigoBox.Text
        Dim dsp As Integer = If(rbtnDisplay.Checked = True, 1, 0)

        '入力チェック
        If nam = "" Then
            MsgBox("漢字氏名を入力して下さい。")
            Return
        ElseIf kana = "" Then
            MsgBox("カナ氏名を入力して下さい。")
            Return
        ElseIf sex = "" Then
            MsgBox("性別を入力して下さい。")
            Return
        ElseIf birth = "" Then
            MsgBox("生年月日を入力して下さい。")
            Return
        ElseIf kaigo <> "" AndAlso IsNumeric(kaigo) = False Then
            MsgBox("介護度を正しく入力して下さい。")
            Return
        ElseIf idBox.Text <> "" AndAlso IsNumeric(idBox.Text) = False Then
            MsgBox("IDは数値を入力してください。")
            Return
        End If

        '空ではない場合、介護度の頭のゼロを取る
        If kaigo <> "" Then
            kaigo = CInt(kaigo).ToString
        End If

        '空ではない場合、IDの頭のゼロを取る
        Dim id As Integer = If(idBox.Text <> "", CInt(idBox.Text), -1)

        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand

        'IDの設定
        If id = -1 Then
            '最新のid+1の値を取得
            SQLCm.CommandText = "SELECT top 1 Id FROM KihonM order by Id desc"
            Cn.Open()
            reader = SQLCm.ExecuteReader()
            If reader.Read() = True Then
                id = reader("Id") + 1
            Else
                id = 1
            End If
            reader.Close()
            Cn.Close()
        End If

        'IDの有無で更新か新規登録
        SQLCm.CommandText = "SELECT Id FROM KihonM where Id=@id"
        SQLCm.Parameters.Clear()
        SQLCm.Parameters.Add("@id", OleDbType.Integer).Value = id
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        If reader.Read() = True Then
            reader.Close()
            '更新
            SQLCm.CommandText = "update KihonM set Nam=@nam, Kana=@kana, Sex=@sex, Birth=@birth, Kaigo=@kaigo, Dsp=@dsp where Id=@id"
            SQLCm.Parameters.Clear()
            SQLCm.Parameters.Add("@nam", OleDbType.Char).Value = nam
            SQLCm.Parameters.Add("@kana", OleDbType.Char).Value = kana
            SQLCm.Parameters.Add("@sex", OleDbType.Char).Value = sex
            SQLCm.Parameters.Add("@birth", OleDbType.Char).Value = birth
            SQLCm.Parameters.Add("@kaigo", OleDbType.Char).Value = kaigo
            SQLCm.Parameters.Add("@dsp", OleDbType.Integer).Value = dsp
            SQLCm.Parameters.Add("@id", OleDbType.Integer).Value = id
            SQLCm.ExecuteNonQuery()
            Cn.Close()
        Else
            reader.Close()
            '新規登録
            SQLCm.CommandText = "insert into KihonM(Id, Nam, Kana, Sex, Birth, Kaigo, Dsp) values (@id, @nam, @kana, @sex, @birth, @kaigo, @dsp)"
            SQLCm.Parameters.Clear()
            SQLCm.Parameters.Add("@id", OleDbType.Integer).Value = id
            SQLCm.Parameters.Add("@nam", OleDbType.Char).Value = nam
            SQLCm.Parameters.Add("@kana", OleDbType.Char).Value = kana
            SQLCm.Parameters.Add("@sex", OleDbType.Char).Value = sex
            SQLCm.Parameters.Add("@birth", OleDbType.Char).Value = birth
            SQLCm.Parameters.Add("@kaigo", OleDbType.Char).Value = kaigo
            SQLCm.Parameters.Add("@dsp", OleDbType.Integer).Value = dsp
            SQLCm.ExecuteNonQuery()
            Cn.Close()
        End If

        '入力テキストクリア
        inputClear()

        '再表示
        displayUserMasterData(dgvUserMaster)

    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        If IsNumeric(idBox.Text) = False Then
            MsgBox("利用者IDに1以上の数値を入力してください。")
            Return
        End If

        Dim targetId As Integer = CInt(idBox.Text)

        '入力されているIDのデータが存在するかチェック
        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        SQLCm.CommandText = "SELECT top 1 * from KihonM where Id=" & targetId
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        If reader.Read() = False Then
            reader.Close()
            MsgBox("登録されていません。")
            Cn.Close()
            Cn.Dispose()
            Return
        Else
            '削除処理
            reader.Close()
            Dim result As DialogResult = MessageBox.Show("削除してよろしいですか？", "Nurse2", MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.Yes Then
                SQLCm.CommandText = "delete from KihonM where Id=" & targetId
                SQLCm.ExecuteNonQuery()
                inputClear()
                Cn.Close()
                Cn.Dispose()

                '再表示
                displayUserMasterData(dgvUserMaster)
            End If
        End If
    End Sub

    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click

    End Sub

    Private Sub dgvUserMaster_CellFormatting(sender As Object, e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dgvUserMaster.CellFormatting
        If dgvUserMaster.Columns(e.ColumnIndex).Name = "Dsp" Then
            If e.Value = "1" Then
                e.Value = "○"
            Else
                e.Value = ""
            End If
            e.FormattingApplied = True
        End If
    End Sub

    Private Sub dgvUserMaster_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvUserMaster.CellMouseClick
        If e.RowIndex = -1 Then
            Return
        End If

        Dim id As String = dgvUserMaster("Id", e.RowIndex).Value
        Dim nam As String = dgvUserMaster("Nam", e.RowIndex).Value
        Dim kana As String = dgvUserMaster("Kana", e.RowIndex).Value
        Dim sex As String = dgvUserMaster("Sex", e.RowIndex).Value
        Dim birth As String = dgvUserMaster("Birth", e.RowIndex).Value
        Dim kaigo As String = dgvUserMaster("Kaigo", e.RowIndex).Value
        Dim dsp As String = dgvUserMaster("Dsp", e.RowIndex).Value

        'テキストボックスへ反映
        idBox.Text = id
        namBox.Text = nam
        kanaBox.Text = kana
        sexBox.Text = sex
        birthYmdBox.setWarakiStr(birth)
        kaigoBox.Text = kaigo
        If dsp = 0 Then
            rbtnNotDisplay.Checked = True
        ElseIf dsp = 1 Then
            rbtnDisplay.Checked = True
        End If

        '選択している行のインデックス保持
        selectedRowIndex = e.RowIndex
    End Sub

    Private Sub dgvUserMaster_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvUserMaster.CellPainting
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

    Private Sub dgvUserMaster_Scroll(sender As Object, e As System.Windows.Forms.ScrollEventArgs) Handles dgvUserMaster.Scroll
        scrollPosition = dgvUserMaster.FirstDisplayedScrollingRowIndex
    End Sub

    Private Sub sexBox_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles sexBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim inputStr As String = sexBox.Text
            If inputStr = "2" Then
                sexBox.Text = "女"
            Else
                sexBox.Text = "男"
            End If
        End If
    End Sub
End Class