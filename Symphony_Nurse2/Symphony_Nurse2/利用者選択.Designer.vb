<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 利用者選択
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dgvUser = New System.Windows.Forms.DataGridView()
        Me.searchIdBox = New System.Windows.Forms.TextBox()
        Me.searchNamBox = New System.Windows.Forms.TextBox()
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvUser
        '
        Me.dgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvUser.DefaultCellStyle = DataGridViewCellStyle1
        Me.dgvUser.Location = New System.Drawing.Point(12, 32)
        Me.dgvUser.Name = "dgvUser"
        Me.dgvUser.RowTemplate.Height = 21
        Me.dgvUser.Size = New System.Drawing.Size(170, 542)
        Me.dgvUser.TabIndex = 2
        '
        'searchIdBox
        '
        Me.searchIdBox.BackColor = System.Drawing.SystemColors.Control
        Me.searchIdBox.ForeColor = System.Drawing.Color.Blue
        Me.searchIdBox.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.searchIdBox.Location = New System.Drawing.Point(12, 7)
        Me.searchIdBox.Name = "searchIdBox"
        Me.searchIdBox.Size = New System.Drawing.Size(44, 19)
        Me.searchIdBox.TabIndex = 0
        '
        'searchNamBox
        '
        Me.searchNamBox.BackColor = System.Drawing.SystemColors.Control
        Me.searchNamBox.ForeColor = System.Drawing.Color.Blue
        Me.searchNamBox.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.searchNamBox.Location = New System.Drawing.Point(62, 7)
        Me.searchNamBox.Name = "searchNamBox"
        Me.searchNamBox.Size = New System.Drawing.Size(120, 19)
        Me.searchNamBox.TabIndex = 1
        '
        '利用者選択
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(198, 584)
        Me.Controls.Add(Me.searchNamBox)
        Me.Controls.Add(Me.searchIdBox)
        Me.Controls.Add(Me.dgvUser)
        Me.Name = "利用者選択"
        Me.Text = "利用者選択"
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents searchIdBox As System.Windows.Forms.TextBox
    Friend WithEvents searchNamBox As System.Windows.Forms.TextBox
End Class
