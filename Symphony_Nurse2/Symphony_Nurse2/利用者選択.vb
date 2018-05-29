Imports System.Data.OleDb

Public Class 利用者選択

    'テキストボックスのマウスダウンイベント制御用
    Private mdFlag As Boolean = False

    Private Sub 利用者選択_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'フォームの設定
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        '位置設定
        Me.Left = 0
        Me.Top = 55

        'DoubleBufferedプロパティをTrue
        TopForm.EnableDoubleBuffering(dgvUser)

        'dgv設定
        With dgvUser
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
            .RowTemplate.Height = 18
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .ShowCellToolTips = False
        End With

        'データ取得、表示
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim dt As New DataTable
        SQLCm.CommandText = "select Id, Nam, Kana, Dsp, Sex, Birth, Kaigo from KihonM order by Kana"
        Adapter.Fill(dt)
        dgvUser.DataSource = dt

        '表示設定
        With dgvUser
            'Idは中央
            .Columns("Id").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            '幅
            .Columns("Id").Width = 45
            .Columns("Nam").Width = 105

            '非表示
            .Columns("Kana").Visible = False
            .Columns("Dsp").Visible = False
            .Columns("Sex").Visible = False
            .Columns("Birth").Visible = False
            .Columns("Kaigo").Visible = False

            'Dsp=1の最初に出現する行を選択させる
            For Each row As DataGridViewRow In dgvUser.Rows
                If Not IsDBNull(row.Cells("Dsp").Value) AndAlso row.Cells("Dsp").Value = 1 Then
                    dgvUser.CurrentCell = dgvUser("Nam", row.Index)
                    Exit For
                End If
            Next

            'Dsp=0の行を非表示にする
            For Each row As DataGridViewRow In dgvUser.Rows
                If IsDBNull(row.Cells("Dsp").Value) OrElse row.Cells("Dsp").Value = 0 Then
                    row.Visible = False
                End If
            Next

        End With

        '検索ID,名前ボックスの設定
        searchIdBox.TextAlign = HorizontalAlignment.Center
        searchNamBox.TextAlign = HorizontalAlignment.Left

        '親フォームに利用者選択の情報があり、且つ、dgvに存在する場合は設定
        Dim parentForm As TopForm = CType(Me.Owner, TopForm)
        Dim nam As String = parentForm.userNam.Text
        Dim searchResult() As DataRow = dt.Select("Nam = '" & nam & "'")
        If searchResult.Length = 1 Then
            searchIdBox.Text = parentForm.userId.Text
            searchNamBox.Text = parentForm.userNam.Text
        Else
            searchIdBox.Text = ""
            searchNamBox.Text = ""

            parentForm.userId.Text = ""
            parentForm.userNam.Text = ""
            parentForm.userKana.Text = ""
            parentForm.userDsp.Text = ""
            parentForm.userBirth.Text = ""
            parentForm.userSex.Text = ""
            parentForm.userKaigo.Text = ""
        End If

        '検索IDボックスにフォーカス
        searchIdBox.Focus()

    End Sub

    Private Function checkDBNullValue(dgvCellValue As Object) As String
        Return If(IsDBNull(dgvCellValue), "", dgvCellValue)
    End Function

    Private Sub dgvUser_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvUser.CellMouseClick
        Dim parentForm As TopForm = CType(Me.Owner, TopForm)
        Dim id As String = checkDBNullValue(dgvUser("Id", e.RowIndex).Value)
        Dim nam As String = checkDBNullValue(dgvUser("Nam", e.RowIndex).Value)
        Dim kana As String = checkDBNullValue(dgvUser("Kana", e.RowIndex).Value)
        Dim dsp As String = checkDBNullValue(dgvUser("Dsp", e.RowIndex).Value)
        Dim sex As String = checkDBNullValue(dgvUser("Sex", e.RowIndex).Value)
        Dim birth As String = checkDBNullValue(dgvUser("Birth", e.RowIndex).Value)
        Dim kaigo As String = checkDBNullValue(dgvUser("Kaigo", e.RowIndex).Value)

        'idとNamを表示
        searchIdBox.Text = id
        searchNamBox.Text = nam

        '親フォームのテキストボックスに情報渡す
        parentForm.userId.Text = id
        parentForm.userNam.Text = nam
        parentForm.userKana.Text = kana
        parentForm.userDsp.Text = dsp
        parentForm.userSex.Text = sex
        parentForm.userBirth.Text = birth
        parentForm.userKaigo.Text = kaigo

    End Sub

    Private Sub searchIdBox_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles searchIdBox.KeyDown
        'エンターキー押下時、IDで検索
        If e.KeyCode = Keys.Enter Then
            Dim parentForm As TopForm = CType(Me.Owner, TopForm)
            Dim id As Integer = If(IsNumeric(searchIdBox.Text), CInt(searchIdBox.Text), -1)
            Dim dt As DataTable = CType(dgvUser.DataSource, DataTable)
            Dim searchResult() As DataRow = dt.Select("Id = " & id)
            If searchResult.Length = 1 Then
                '入力されたIDが存在する場合
                searchNamBox.Text = searchResult(0).Item("Nam")

                '親フォームに反映
                parentForm.userId.Text = id
                parentForm.userNam.Text = searchResult(0).Item("Nam")
                parentForm.userKana.Text = searchResult(0).Item("Kana")
                parentForm.userDsp.Text = searchResult(0).Item("Dsp")
                parentForm.userSex.Text = searchResult(0).Item("Sex")
                parentForm.userBirth.Text = searchResult(0).Item("Birth")
                parentForm.userKaigo.Text = searchResult(0).Item("Kaigo")
            Else
                '存在しない場合
                searchIdBox.Text = ""
                searchNamBox.Text = ""

                '親フォームに反映
                parentForm.userId.Text = ""
                parentForm.userNam.Text = ""
                parentForm.userKana.Text = ""
                parentForm.userDsp.Text = ""
                parentForm.userSex.Text = ""
                parentForm.userBirth.Text = ""
                parentForm.userKaigo.Text = ""
            End If
        End If
    End Sub

    Private Sub searchNamBox_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles searchNamBox.KeyDown
        'エンターキー押下時、氏名で検索
        If e.KeyCode = Keys.Enter Then
            Dim parentForm As TopForm = CType(Me.Owner, TopForm)
            Dim nam As String = searchNamBox.Text
            Dim dt As DataTable = CType(dgvUser.DataSource, DataTable)
            Dim searchResult() As DataRow = dt.Select("Nam = '" & nam & "'")
            If searchResult.Length = 1 Then
                '入力された氏名が存在する場合
                searchIdBox.Text = searchResult(0).Item("Id")

                '親フォームに反映
                parentForm.userId.Text = searchResult(0).Item("Id")
                parentForm.userNam.Text = nam
                parentForm.userKana.Text = searchResult(0).Item("Kana")
                parentForm.userDsp.Text = searchResult(0).Item("Dsp")
                parentForm.userSex.Text = searchResult(0).Item("Sex")
                parentForm.userBirth.Text = searchResult(0).Item("Birth")
                parentForm.userKaigo.Text = searchResult(0).Item("Kaigo")
            Else
                '存在しない場合
                searchIdBox.Text = ""
                searchNamBox.Text = ""

                '親フォームに反映
                parentForm.userId.Text = ""
                parentForm.userNam.Text = ""
                parentForm.userKana.Text = ""
                parentForm.userDsp.Text = ""
                parentForm.userSex.Text = ""
                parentForm.userBirth.Text = ""
                parentForm.userKaigo.Text = ""
            End If
        End If
    End Sub

    Private Sub textBox_Enter(sender As Object, e As System.EventArgs) Handles searchIdBox.Enter, searchNamBox.Enter
        Dim tb As TextBox = CType(sender, TextBox)
        tb.SelectAll()
        mdFlag = True
    End Sub

    Private Sub textBox_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles searchIdBox.MouseDown, searchNamBox.MouseDown
        If mdFlag = True Then
            Dim tb As TextBox = CType(sender, TextBox)
            tb.SelectAll()
            mdFlag = False
        End If
    End Sub
End Class