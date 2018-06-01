<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 印刷設定
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.rbtnPrint = New System.Windows.Forms.RadioButton()
        Me.rbtnPreview = New System.Windows.Forms.RadioButton()
        Me.SuspendLayout()
        '
        'rbtnPrint
        '
        Me.rbtnPrint.AutoSize = True
        Me.rbtnPrint.Location = New System.Drawing.Point(103, 29)
        Me.rbtnPrint.Name = "rbtnPrint"
        Me.rbtnPrint.Size = New System.Drawing.Size(47, 16)
        Me.rbtnPrint.TabIndex = 0
        Me.rbtnPrint.Text = "印刷"
        Me.rbtnPrint.UseVisualStyleBackColor = True
        '
        'rbtnPreview
        '
        Me.rbtnPreview.AutoSize = True
        Me.rbtnPreview.Checked = True
        Me.rbtnPreview.Location = New System.Drawing.Point(24, 29)
        Me.rbtnPreview.Name = "rbtnPreview"
        Me.rbtnPreview.Size = New System.Drawing.Size(67, 16)
        Me.rbtnPreview.TabIndex = 1
        Me.rbtnPreview.TabStop = True
        Me.rbtnPreview.Text = "プレビュー"
        Me.rbtnPreview.UseVisualStyleBackColor = True
        '
        '印刷設定
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(180, 77)
        Me.Controls.Add(Me.rbtnPreview)
        Me.Controls.Add(Me.rbtnPrint)
        Me.Name = "印刷設定"
        Me.Text = "印刷設定"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rbtnPrint As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnPreview As System.Windows.Forms.RadioButton
End Class
