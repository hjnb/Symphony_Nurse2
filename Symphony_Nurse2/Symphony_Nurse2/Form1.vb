Public Class TopForm

    Private Sub TopForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 終了ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 終了ToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub 利用者選択ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 利用者選択ToolStripMenuItem.Click

    End Sub
End Class
