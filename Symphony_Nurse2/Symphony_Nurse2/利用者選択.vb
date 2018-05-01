Imports System.Data.OleDb

Public Class 利用者選択

    Private Sub 利用者選択_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        searchIdBox.Focus()

        '入力されたid, Namでエンターで検索する部分作成

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
        End With

        'データ取得、表示
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        SQLCm.CommandText = "select Id, Nam, Kana, Dsp, Sex, Birth, Kaigo from KihonM where Dsp=1 order by Kana"
        Adapter.Fill(Table)
        dgvUser.DataSource = Table

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

        End With

    End Sub

    Private Sub dgvUser_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvUser.CellMouseClick
        Dim parentForm As TopForm = CType(Me.Owner, TopForm)
        Dim id As String = dgvUser("Id", e.RowIndex).Value
        Dim nam As String = dgvUser("Nam", e.RowIndex).Value
        Dim kana As String = dgvUser("Kana", e.RowIndex).Value
        Dim dsp As String = dgvUser("Dsp", e.RowIndex).Value
        Dim sex As String = dgvUser("Sex", e.RowIndex).Value
        Dim birth As String = dgvUser("Birth", e.RowIndex).Value
        Dim kaigo As String = dgvUser("Kaigo", e.RowIndex).Value

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
End Class