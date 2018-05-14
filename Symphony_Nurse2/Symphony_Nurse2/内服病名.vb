Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Public Class 内服病名

    Private Sub 内服病名_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim forward As Boolean = e.Modifiers <> Keys.Shift
            Me.SelectNextControl(Me.ActiveControl, forward, True, True, True)
            e.Handled = True
        End If
    End Sub

    Private Sub 内服病名_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
        Me.DesktopLocation = New Point(0, 50)

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        SQLCm.CommandText = "select * from Unit order by Gyo"
        Adapter.Fill(Table)
        DataGridView3.DataSource = Table

        With DataGridView1
            .Columns.Add("RoomNo", "　")
            .Columns.Add("Nam", "氏名")
            .Columns(0).Width = 30
            .Columns(1).Width = 100
            .RowTemplate.Height = 18
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
        End With

        '名前をセット
        For n As Integer = 1 To 10
            For i As Integer = 0 To 9
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i + 10 * (n - 1)).Cells(1).Value = DataGridView3(3 * n - 1, i).Value
            Next
        Next

        '部屋番号をセット
        Dim Ha As Integer = 1
        Dim Hb As Integer = 1
        Dim Hc As Integer = 1
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            For C As Integer = Hc To 2       '本館とアネックス分
                For A As Integer = Ha To 6       '部屋番号の頭
                    If A = 4 Then
                        '何もしない
                        Ha = Ha + 1
                    ElseIf A = 6 Then
                        For B As Integer = Hb To 12      '部屋番号の後ろ
                            If B = 4 Then
                                '何もしない
                                Hb = Hb + 1
                                'GoTo line1
                            ElseIf B = 9 Then
                                '何もしない
                                Hb = Hb + 1
                                'GoTo line1
                            ElseIf B < 10 Then      'Bが一桁の時
                                DataGridView1.Rows(i).Cells(0).Value = A & "0" & B
                                Hb = Hb + 1
                                GoTo line1
                            ElseIf B = 12 Then
                                DataGridView1.Rows(i).Cells(0).Value = A & B
                                'Ha = Ha + 1
                                Hb = 1
                                Ha = 1
                                Hc = Hc + 1
                                GoTo line1
                            ElseIf B > 9 Then       'Bが二桁の時
                                DataGridView1.Rows(i).Cells(0).Value = A & B
                                Hb = Hb + 1
                                GoTo line1
                            End If
                        Next
                    Else
                        For B As Integer = Hb To 12      '部屋番号の後ろ
                            If B = 4 Then
                                '何もしない
                                Hb = Hb + 1
                                'GoTo line1
                            ElseIf B = 9 Then
                                '何もしない
                                Hb = Hb + 1
                                'GoTo line1
                            ElseIf B < 10 Then      'Bが一桁の時
                                DataGridView1.Rows(i).Cells(0).Value = A & "0" & B
                                Hb = Hb + 1
                                GoTo line1
                            ElseIf B = 12 Then
                                DataGridView1.Rows(i).Cells(0).Value = A & B
                                Ha = Ha + 1
                                Hb = 1
                                GoTo line1
                            ElseIf B > 9 Then       'Bが二桁の時
                                DataGridView1.Rows(i).Cells(0).Value = A & B
                                Hb = Hb + 1
                                GoTo line1
                            End If
                        Next
                    End If
                Next
            Next
line1:
        Next

        'ユニット名の追加と設定
        With DataGridView1
            .Rows.Insert(0)
            .Rows(0).Cells(1).Value = "空"
            .Rows(0).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(11)
            .Rows(11).Cells(1).Value = "森"
            .Rows(11).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(22)
            .Rows(22).Cells(1).Value = "星"
            .Rows(22).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(33)
            .Rows(33).Cells(1).Value = "月"
            .Rows(33).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(44)
            .Rows(44).Cells(1).Value = "花"
            .Rows(44).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(55)
            .Rows(55).Cells(1).Value = "丘"
            .Rows(55).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(66)
            .Rows(66).Cells(1).Value = "虹"
            .Rows(66).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(77)
            .Rows(77).Cells(1).Value = "光"
            .Rows(77).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(88)
            .Rows(88).Cells(1).Value = "雪"
            .Rows(88).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            .Rows.Insert(99)
            .Rows(99).Cells(1).Value = "風"
            .Rows(99).Cells(1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
        End With

        lblID.Text = ""

        KeyPreview = True
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        btnKuria.PerformClick()
        DataGridView4.Columns.Clear()
        Dim s As Integer = DataGridView1.CurrentRow.Index
        lblName.Text = DataGridView1(1, s).Value
        If DataGridView1(0, s).Value <> "" Then         '部屋番号がある行を選択したとき
            If s < 11 Then
                lblHeya.Text = "空　" & DataGridView1(0, s).Value
                For i As Integer = 0 To DataGridView3.Rows.Count - 1
                    If lblName.Text = DataGridView3(2, i).Value Then
                        lblID.Text = DataGridView3(1, i).Value
                        Exit For
                    End If
                Next
            ElseIf s < 22 Then
                lblHeya.Text = "森　" & DataGridView1(0, s).Value
                For i As Integer = 0 To DataGridView3.Rows.Count - 1
                    If lblName.Text = DataGridView3(5, i).Value Then
                        lblID.Text = DataGridView3(4, i).Value
                        Exit For
                    End If
                Next
            ElseIf s < 33 Then
                lblHeya.Text = "星　" & DataGridView1(0, s).Value
                For i As Integer = 0 To DataGridView3.Rows.Count - 1
                    If lblName.Text = DataGridView3(8, i).Value Then
                        lblID.Text = DataGridView3(7, i).Value
                        Exit For
                    End If
                Next
            ElseIf s < 44 Then
                lblHeya.Text = "月　" & DataGridView1(0, s).Value
                For i As Integer = 0 To DataGridView3.Rows.Count - 1
                    If lblName.Text = DataGridView3(11, i).Value Then
                        lblID.Text = DataGridView3(10, i).Value
                        Exit For
                    End If
                Next
            ElseIf s < 55 Then
                lblHeya.Text = "花　" & DataGridView1(0, s).Value
                For i As Integer = 0 To DataGridView3.Rows.Count - 1
                    If lblName.Text = DataGridView3(14, i).Value Then
                        lblID.Text = DataGridView3(13, i).Value
                        Exit For
                    End If
                Next
            ElseIf s < 66 Then
                lblHeya.Text = "丘　" & DataGridView1(0, s).Value
                For i As Integer = 0 To DataGridView3.Rows.Count - 1
                    If lblName.Text = DataGridView3(17, i).Value Then
                        lblID.Text = DataGridView3(16, i).Value
                        Exit For
                    End If
                Next
            ElseIf s < 77 Then
                lblHeya.Text = "虹　" & DataGridView1(0, s).Value
                For i As Integer = 0 To DataGridView3.Rows.Count - 1
                    If lblName.Text = DataGridView3(20, i).Value Then
                        lblID.Text = DataGridView3(19, i).Value
                        Exit For
                    End If
                Next
            ElseIf s < 88 Then
                lblHeya.Text = "光　" & DataGridView1(0, s).Value
                For i As Integer = 0 To DataGridView3.Rows.Count - 1
                    If lblName.Text = DataGridView3(23, i).Value Then
                        lblID.Text = DataGridView3(22, i).Value
                        Exit For
                    End If
                Next
            ElseIf s < 99 Then
                lblHeya.Text = "雪　" & DataGridView1(0, s).Value
                For i As Integer = 0 To DataGridView3.Rows.Count - 1
                    If lblName.Text = DataGridView3(26, i).Value Then
                        lblID.Text = DataGridView3(25, i).Value
                        Exit For
                    End If
                Next
            ElseIf s < 110 Then
                lblHeya.Text = "風　" & DataGridView1(0, s).Value
                For i As Integer = 0 To DataGridView3.Rows.Count - 1
                    If lblName.Text = DataGridView3(29, i).Value Then
                        lblID.Text = DataGridView3(28, i).Value
                        Exit For
                    End If
                Next
            End If
        Else         '部屋番号がない行を選択したとき
            lblHeya.Text = ""
            lblID.Text = ""
            lblName.Text = ""
            '何もしない
        End If

        If lblID.Text <> "" Then
            Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
            Dim SQLCm As OleDbCommand = Cn.CreateCommand
            Dim Adapter As New OleDbDataAdapter(SQLCm)
            Dim Table As New DataTable
            SQLCm.CommandText = "select * from BSht WHERE Id = " & lblID.Text & " order by Gyo"
            Adapter.Fill(Table)
            DataGridView4.DataSource = Table

            'データがあったら表示
            If DataGridView4.Rows.Count > 15 Then
                For Gyo As Integer = 1 To 16
                    If Gyo < 5 Then
                        If IsDBNull(DataGridView4(5, Gyo - 1).Value) = False Then
                            Controls("txtByoumei" & Gyo).Text = DataGridView4(5, Gyo - 1).Value
                        End If
                        If IsDBNull(DataGridView4(10, Gyo - 1).Value) = False Then
                            Controls("txtTokki" & Gyo).Text = DataGridView4(10, Gyo - 1).Value
                        End If
                    End If
                    Controls("txtNaihuku" & Gyo).Text = DataGridView4(6, Gyo - 1).Value
                    Controls("txtRyou" & Gyo).Text = DataGridView4(7, Gyo - 1).Value
                    Controls("txtKatati" & Gyo).Text = DataGridView4(8, Gyo - 1).Value
                    Controls("txtJikann" & Gyo).Text = DataGridView4(9, Gyo - 1).Value
                Next
                YmdBox1.setADStr(DataGridView4(11, 0).Value)
            End If

        End If


    End Sub

    Private Sub btnKuria_Click(sender As System.Object, e As System.EventArgs) Handles btnKuria.Click
        lblHeya.Text = ""
        lblID.Text = ""
        lblName.Text = ""
        For i As Integer = 1 To 4
            Controls("txtByoumei" & i).Text = ""
            Controls("txtTokki" & i).Text = ""
        Next
        For i As Integer = 1 To 16
            Controls("txtNaihuku" & i).Text = ""
            Controls("txtRyou" & i).Text = ""
            Controls("txtKatati" & i).Text = ""
            Controls("txtJikann" & i).Text = ""
        Next
        YmdBox1.clearText()
    End Sub

    Private Sub btnTouroku_Click(sender As System.Object, e As System.EventArgs) Handles btnTouroku.Click
        If YmdBox1.getADStr() = "" Then
            MsgBox("日付を入力してください。")
            Return
        End If

        If DataGridView4.Rows.Count <> 0 Then
            'データがあったら変更
            Hennkou()
            btnKousinn.PerformClick()
        Else
            'データがなかったら新規追加
            Tuika()
            btnKousinn.PerformClick()
        End If

    End Sub
    Private Sub Hennkou()
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim SQL As String = ""
        SQL = "DELETE FROM BSht WHERE (Id = " & lblID.Text & ") AND (Ymd ='" & YmdBox1.getADStr() & "')"
        SQLCm.CommandText = SQL
        Cn.Open()
        SQLCm.ExecuteNonQuery()
        Cn.Close()

        SQLCm.Dispose()
        Cn.Dispose()

        Tuika()

    End Sub
    Private Sub Tuika()
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Table As DataTable = DirectCast(DataGridView4.DataSource, DataTable)
        Dim Row As DataRow = Table.NewRow

        Cn.Open()
        For Gyo As Integer = 1 To 16
            If Gyo < 5 Then
                Row("Bmei") = Controls("txtByoumei" & Gyo).Text
                Row("Tokki") = Controls("txtTokki" & Gyo).Text
            Else
                Row("Bmei") = ""
                Row("Tokki") = ""
            End If
            Row("Id") = lblID.Text
            Row("Nam") = lblName.Text
            Row("Unt") = Strings.Left(lblHeya.Text, 1)
            Row("Rm") = Strings.Right(lblHeya.Text, 3)
            Row("Gyo") = Gyo
            Row("Md1") = Controls("txtNaihuku" & Gyo).Text
            Row("Md2") = Controls("txtRyou" & Gyo).Text
            Row("Md3") = Controls("txtKatati" & Gyo).Text
            Row("Md4") = Controls("txtJikann" & Gyo).Text
            Row("Ymd") = YmdBox1.getADStr()
            Dim SQL As String = ""
            SQL = "INSERT INTO BSht (Id, Nam, Unt, Rm, Gyo, Bmei, Md1, Md2, Md3, Md4, Tokki, Ymd) VALUES ("
            SQL &= Row("Id") & ", "
            SQL &= "'" & Row("Nam") & "', "
            SQL &= "'" & Row("Unt") & "', "
            SQL &= "'" & Row("Rm") & "', "
            SQL &= "'" & Row("Gyo") & "', "
            SQL &= "'" & Row("Bmei") & "', "
            SQL &= "'" & Row("Md1") & "', "
            SQL &= "'" & Row("Md2") & "', "
            SQL &= "'" & Row("Md3") & "', "
            SQL &= "'" & Row("Md4") & "', "
            SQL &= "'" & Row("Tokki") & "', "
            SQL &= "'" & Row("Ymd") & "' "
            SQL &= ")"

            SQLCm.CommandText = SQL
            SQLCm.ExecuteNonQuery()
        Next

        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()
    End Sub

    Private Sub btnInnsatu_Click(sender As System.Object, e As System.EventArgs) Handles btnInnsatu.Click

    End Sub

    Private Sub btnKousinn_Click(sender As System.Object, e As System.EventArgs) Handles btnKousinn.Click
        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        SQLCm.CommandText = "select * from BSht WHERE Id = " & lblID.Text & " order by Gyo"
        Adapter.Fill(Table)
        DataGridView4.DataSource = Table
    End Sub

    Private Sub rbnSubete_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnSubete.CheckedChanged

    End Sub

    Private Sub rbnKigou_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnKigou.CheckedChanged

    End Sub

    Private Sub rbnA_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnA.CheckedChanged

    End Sub

    Private Sub rbnK_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnK.CheckedChanged

    End Sub

    Private Sub rbnS_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnS.CheckedChanged

    End Sub

    Private Sub rbnT_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnT.CheckedChanged

    End Sub

    Private Sub rbnN_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnN.CheckedChanged

    End Sub

    Private Sub rbnH_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnH.CheckedChanged

    End Sub

    Private Sub rbnM_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnM.CheckedChanged

    End Sub

    Private Sub rbnY_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnY.CheckedChanged

    End Sub

    Private Sub rbnR_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnR.CheckedChanged

    End Sub

    Private Sub rbnW_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbnW.CheckedChanged

    End Sub


End Class