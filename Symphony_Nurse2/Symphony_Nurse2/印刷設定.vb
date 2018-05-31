Public Class 印刷設定

    Private p As TopForm

    Private Sub 印刷設定_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Left = 210
        Me.Top = 55
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        p = CType(Me.Owner, TopForm)
        If p.rbtnPreview.Checked = True Then
            Me.rbtnPreview.Checked = True
        Else
            Me.rbtnPrint.Checked = True
        End If
    End Sub

    Private Sub rbtnPreview_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles rbtnPreview.MouseClick
        p.rbtnPreview.Checked = True
    End Sub

    Private Sub rbtnPrint_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles rbtnPrint.MouseClick
        p.rbtnPrint.Checked = True
    End Sub
End Class