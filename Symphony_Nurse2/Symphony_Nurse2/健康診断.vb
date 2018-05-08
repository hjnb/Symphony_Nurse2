Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Public Class 健康診断

    Private Sub 健康診断_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
        Me.DesktopLocation = New Point(150, 50)

        Dim Cn As New OleDbConnection(TopForm.DB_Nurse2)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        SQLCm.CommandText = "select Id, Nam, Kana, Dsp, Sex, Birth, Kaigo from KihonM where Dsp=1 order by Kana"
        Adapter.Fill(Table)
        DataGridView3.DataSource = Table

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