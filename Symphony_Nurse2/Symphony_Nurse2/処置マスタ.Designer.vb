<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 処置マスタ
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbCategory1 = New System.Windows.Forms.ComboBox()
        Me.dspBox = New System.Windows.Forms.TextBox()
        Me.category2Box = New System.Windows.Forms.TextBox()
        Me.category3Box = New System.Windows.Forms.TextBox()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.dgvTreatingMaster = New System.Windows.Forms.DataGridView()
        CType(Me.dgvTreatingMaster, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(28, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "分類1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(28, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 12)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "表示順"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(28, 86)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(35, 12)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "分類2"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(28, 113)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(35, 12)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "分類3"
        '
        'cmbCategory1
        '
        Me.cmbCategory1.FormattingEnabled = True
        Me.cmbCategory1.Location = New System.Drawing.Point(86, 27)
        Me.cmbCategory1.Name = "cmbCategory1"
        Me.cmbCategory1.Size = New System.Drawing.Size(100, 20)
        Me.cmbCategory1.TabIndex = 4
        '
        'dspBox
        '
        Me.dspBox.Location = New System.Drawing.Point(86, 56)
        Me.dspBox.Name = "dspBox"
        Me.dspBox.Size = New System.Drawing.Size(48, 19)
        Me.dspBox.TabIndex = 5
        '
        'category2Box
        '
        Me.category2Box.Location = New System.Drawing.Point(86, 83)
        Me.category2Box.Name = "category2Box"
        Me.category2Box.Size = New System.Drawing.Size(239, 19)
        Me.category2Box.TabIndex = 6
        '
        'category3Box
        '
        Me.category3Box.Location = New System.Drawing.Point(86, 110)
        Me.category3Box.Name = "category3Box"
        Me.category3Box.Size = New System.Drawing.Size(48, 19)
        Me.category3Box.TabIndex = 7
        '
        'btnRegist
        '
        Me.btnRegist.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnRegist.Location = New System.Drawing.Point(408, 44)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(80, 33)
        Me.btnRegist.TabIndex = 8
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(408, 93)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 33)
        Me.btnDelete.TabIndex = 9
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'dgvTreatingMaster
        '
        Me.dgvTreatingMaster.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTreatingMaster.Location = New System.Drawing.Point(27, 156)
        Me.dgvTreatingMaster.Name = "dgvTreatingMaster"
        Me.dgvTreatingMaster.RowTemplate.Height = 21
        Me.dgvTreatingMaster.Size = New System.Drawing.Size(469, 322)
        Me.dgvTreatingMaster.TabIndex = 10
        '
        '処置マスタ
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(525, 500)
        Me.Controls.Add(Me.dgvTreatingMaster)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.category3Box)
        Me.Controls.Add(Me.category2Box)
        Me.Controls.Add(Me.dspBox)
        Me.Controls.Add(Me.cmbCategory1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "処置マスタ"
        Me.Text = "処置マスタ"
        CType(Me.dgvTreatingMaster, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbCategory1 As System.Windows.Forms.ComboBox
    Friend WithEvents dspBox As System.Windows.Forms.TextBox
    Friend WithEvents category2Box As System.Windows.Forms.TextBox
    Friend WithEvents category3Box As System.Windows.Forms.TextBox
    Friend WithEvents btnRegist As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents dgvTreatingMaster As System.Windows.Forms.DataGridView
End Class
