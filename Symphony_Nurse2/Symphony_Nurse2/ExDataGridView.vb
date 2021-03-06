﻿Public Class ExDataGridView
    Inherits DataGridView

    Private inputFlg As Boolean = True

    Protected Overrides Function ProcessKeyEventArgs(ByRef m As System.Windows.Forms.Message) As Boolean
        Dim code As Integer = CInt(m.WParam)
        If inputFlg Then
            If code = Keys.Left OrElse code = Keys.Right OrElse code = Keys.Up OrElse code = Keys.Down OrElse code = Keys.Enter OrElse (Keys.NumPad0 <= code AndAlso code <= Keys.NumPad9) OrElse (Keys.D0 <= code AndAlso code <= Keys.D9) OrElse code = Keys.Decimal OrElse code = Keys.OemPeriod Then
                Return MyBase.ProcessKeyEventArgs(m)
            Else
                m.WParam = Keys.Decimal
                Return MyBase.ProcessKeyEventArgs(m)
            End If
        Else
            Return False
        End If
    End Function

    Protected Overrides Function ProcessDataGridViewKey(e As System.Windows.Forms.KeyEventArgs) As Boolean
        Dim tb As DataGridViewTextBoxEditingControl = CType(Me.EditingControl, DataGridViewTextBoxEditingControl)
        If Not IsNothing(tb) AndAlso ((e.KeyCode = Keys.Left AndAlso tb.SelectionStart = 0) OrElse (e.KeyCode = Keys.Right AndAlso tb.SelectionStart = tb.TextLength)) Then
            Return False
        Else
            Return MyBase.ProcessDataGridViewKey(e)
        End If
    End Function

    Protected Overrides Sub OnCellBeginEdit(e As System.Windows.Forms.DataGridViewCellCancelEventArgs)
        inputFlg = False
        MyBase.OnCellBeginEdit(e)
    End Sub

    Protected Overrides Sub OnCellEnter(e As System.Windows.Forms.DataGridViewCellEventArgs)
        inputFlg = True
        MyBase.OnCellEnter(e)
    End Sub

End Class
