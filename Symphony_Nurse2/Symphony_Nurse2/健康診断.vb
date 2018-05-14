Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Public Class 健康診断

    Private Sub 健康診断_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
        Me.DesktopLocation = New Point(180, 50)

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

    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick

        Dim r As Integer = DataGridView1.CurrentRow.Index
        lblName.Text = DataGridView1(1, r).Value
        For i As Integer = 0 To DataGridView3.RowCount - 1

        Next

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable



        'SQLCm.CommandText = "SELECT Cod, Ymd AS 日付, Kbn, Hm AS 時刻, Gyo, Text AS 訴え・観察・判断・処置, Sign FROM KirN" & page & " WHERE Cod = " & a & " ORDER BY Ymd DeSC, Hm DESC, Gyo ASC"
        Adapter.Fill(Table)
        DataGridView1.DataSource = Table
    End Sub
End Class