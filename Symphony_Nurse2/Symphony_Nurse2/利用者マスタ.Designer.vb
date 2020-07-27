<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 利用者マスタ
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
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.rbtnDisplay = New System.Windows.Forms.RadioButton()
        Me.rbtnNotDisplay = New System.Windows.Forms.RadioButton()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.idBox = New System.Windows.Forms.TextBox()
        Me.namBox = New System.Windows.Forms.TextBox()
        Me.kanaBox = New System.Windows.Forms.TextBox()
        Me.sexBox = New System.Windows.Forms.TextBox()
        Me.kaigoBox = New System.Windows.Forms.TextBox()
        Me.dgvUserMaster = New System.Windows.Forms.DataGridView()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.birthYmdBox = New ADBox2.ADBox2()
        CType(Me.dgvUserMaster, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(23, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 14)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "利用者ID"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(23, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 14)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "漢字氏名"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(272, 89)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 14)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "生年月日"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(23, 120)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 14)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "要介護"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(23, 91)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(35, 14)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "性別"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(272, 61)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 14)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "カナ氏名"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(272, 117)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(35, 14)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "表示"
        '
        'rbtnDisplay
        '
        Me.rbtnDisplay.AutoSize = True
        Me.rbtnDisplay.Checked = True
        Me.rbtnDisplay.Location = New System.Drawing.Point(340, 117)
        Me.rbtnDisplay.Name = "rbtnDisplay"
        Me.rbtnDisplay.Size = New System.Drawing.Size(47, 16)
        Me.rbtnDisplay.TabIndex = 7
        Me.rbtnDisplay.TabStop = True
        Me.rbtnDisplay.Text = "表示"
        Me.rbtnDisplay.UseVisualStyleBackColor = True
        '
        'rbtnNotDisplay
        '
        Me.rbtnNotDisplay.AutoSize = True
        Me.rbtnNotDisplay.Location = New System.Drawing.Point(408, 117)
        Me.rbtnNotDisplay.Name = "rbtnNotDisplay"
        Me.rbtnNotDisplay.Size = New System.Drawing.Size(59, 16)
        Me.rbtnNotDisplay.TabIndex = 8
        Me.rbtnNotDisplay.Text = "非表示"
        Me.rbtnNotDisplay.UseVisualStyleBackColor = True
        '
        'btnRegist
        '
        Me.btnRegist.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnRegist.Location = New System.Drawing.Point(376, 148)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(78, 32)
        Me.btnRegist.TabIndex = 6
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(530, 148)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(78, 32)
        Me.btnPrint.TabIndex = 10
        Me.btnPrint.Text = "印刷"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(453, 148)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(78, 32)
        Me.btnDelete.TabIndex = 11
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'idBox
        '
        Me.idBox.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.idBox.Location = New System.Drawing.Point(90, 28)
        Me.idBox.Name = "idBox"
        Me.idBox.Size = New System.Drawing.Size(69, 19)
        Me.idBox.TabIndex = 0
        '
        'namBox
        '
        Me.namBox.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.namBox.Location = New System.Drawing.Point(90, 59)
        Me.namBox.Name = "namBox"
        Me.namBox.Size = New System.Drawing.Size(140, 19)
        Me.namBox.TabIndex = 1
        '
        'kanaBox
        '
        Me.kanaBox.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf
        Me.kanaBox.Location = New System.Drawing.Point(343, 59)
        Me.kanaBox.Name = "kanaBox"
        Me.kanaBox.Size = New System.Drawing.Size(140, 19)
        Me.kanaBox.TabIndex = 2
        '
        'sexBox
        '
        Me.sexBox.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.sexBox.Location = New System.Drawing.Point(90, 89)
        Me.sexBox.Name = "sexBox"
        Me.sexBox.Size = New System.Drawing.Size(50, 19)
        Me.sexBox.TabIndex = 3
        '
        'kaigoBox
        '
        Me.kaigoBox.Location = New System.Drawing.Point(90, 117)
        Me.kaigoBox.Name = "kaigoBox"
        Me.kaigoBox.Size = New System.Drawing.Size(50, 19)
        Me.kaigoBox.TabIndex = 5
        '
        'dgvUserMaster
        '
        Me.dgvUserMaster.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvUserMaster.Location = New System.Drawing.Point(26, 196)
        Me.dgvUserMaster.Name = "dgvUserMaster"
        Me.dgvUserMaster.RowTemplate.Height = 21
        Me.dgvUserMaster.Size = New System.Drawing.Size(605, 385)
        Me.dgvUserMaster.TabIndex = 12
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(151, 92)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(86, 14)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "（1：男　2：女）"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(63, 173)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(161, 12)
        Me.Label9.TabIndex = 14
        Me.Label9.Text = "ﾀﾞﾌﾞﾙｸﾘｯｸした項目名で並べます"
        '
        'birthYmdBox
        '
        Me.birthYmdBox.dateText = ""
        Me.birthYmdBox.Location = New System.Drawing.Point(343, 79)
        Me.birthYmdBox.Mode = 2
        Me.birthYmdBox.monthText = ""
        Me.birthYmdBox.Name = "birthYmdBox"
        Me.birthYmdBox.Size = New System.Drawing.Size(105, 32)
        Me.birthYmdBox.TabIndex = 4
        Me.birthYmdBox.textReadOnly = False
        Me.birthYmdBox.yearText = ""
        '
        '利用者マスタ
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(662, 599)
        Me.Controls.Add(Me.birthYmdBox)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.dgvUserMaster)
        Me.Controls.Add(Me.kaigoBox)
        Me.Controls.Add(Me.sexBox)
        Me.Controls.Add(Me.kanaBox)
        Me.Controls.Add(Me.namBox)
        Me.Controls.Add(Me.idBox)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.rbtnNotDisplay)
        Me.Controls.Add(Me.rbtnDisplay)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.KeyPreview = True
        Me.Name = "利用者マスタ"
        Me.Text = "利用者マスタ"
        CType(Me.dgvUserMaster, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents rbtnDisplay As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnNotDisplay As System.Windows.Forms.RadioButton
    Friend WithEvents btnRegist As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents idBox As System.Windows.Forms.TextBox
    Friend WithEvents namBox As System.Windows.Forms.TextBox
    Friend WithEvents kanaBox As System.Windows.Forms.TextBox
    Friend WithEvents sexBox As System.Windows.Forms.TextBox
    Friend WithEvents kaigoBox As System.Windows.Forms.TextBox
    Friend WithEvents dgvUserMaster As System.Windows.Forms.DataGridView
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents birthYmdBox As ADBox2.ADBox2
End Class
