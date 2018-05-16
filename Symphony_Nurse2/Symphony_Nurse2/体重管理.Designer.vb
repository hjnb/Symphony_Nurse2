<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 体重管理
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
        Me.dgvUser = New System.Windows.Forms.DataGridView()
        Me.selectUserLabel = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dgvUnitLeft = New System.Windows.Forms.DataGridView()
        Me.dgvUnitRight = New System.Windows.Forms.DataGridView()
        Me.rbtnHonkan = New System.Windows.Forms.RadioButton()
        Me.rbtnAnex = New System.Windows.Forms.RadioButton()
        Me.btnPrevNam = New System.Windows.Forms.Button()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.btnPrevCmpr = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.labelUnit1 = New System.Windows.Forms.Label()
        Me.labelUnit3 = New System.Windows.Forms.Label()
        Me.labelUnit2 = New System.Windows.Forms.Label()
        Me.labelUnit5 = New System.Windows.Forms.Label()
        Me.labelUnit6 = New System.Windows.Forms.Label()
        Me.labelUnit4 = New System.Windows.Forms.Label()
        Me.ymdBoxUnit1 = New ymdBox.ymdBox()
        Me.ymdBoxUnit3 = New ymdBox.ymdBox()
        Me.ymdBoxUnit2 = New ymdBox.ymdBox()
        Me.ymdBoxUnit5 = New ymdBox.ymdBox()
        Me.ymdBoxUnit6 = New ymdBox.ymdBox()
        Me.ymdBoxUnit4 = New ymdBox.ymdBox()
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvUnitLeft, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvUnitRight, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvUser
        '
        Me.dgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvUser.Location = New System.Drawing.Point(12, 82)
        Me.dgvUser.Name = "dgvUser"
        Me.dgvUser.RowTemplate.Height = 21
        Me.dgvUser.Size = New System.Drawing.Size(214, 472)
        Me.dgvUser.TabIndex = 1
        '
        'selectUserLabel
        '
        Me.selectUserLabel.AutoSize = True
        Me.selectUserLabel.Font = New System.Drawing.Font("MS UI Gothic", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.selectUserLabel.ForeColor = System.Drawing.Color.MediumBlue
        Me.selectUserLabel.Location = New System.Drawing.Point(22, 36)
        Me.selectUserLabel.Name = "selectUserLabel"
        Me.selectUserLabel.Size = New System.Drawing.Size(0, 27)
        Me.selectUserLabel.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(14, 564)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(180, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "挿入：　ﾏｽﾀ選択後、氏名クリック"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Blue
        Me.Label2.Location = New System.Drawing.Point(14, 584)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(149, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "削除：　室号をダブルクリック"
        '
        'dgvUnitLeft
        '
        Me.dgvUnitLeft.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvUnitLeft.Location = New System.Drawing.Point(241, 44)
        Me.dgvUnitLeft.Name = "dgvUnitLeft"
        Me.dgvUnitLeft.RowTemplate.Height = 21
        Me.dgvUnitLeft.Size = New System.Drawing.Size(273, 502)
        Me.dgvUnitLeft.TabIndex = 5
        '
        'dgvUnitRight
        '
        Me.dgvUnitRight.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvUnitRight.Location = New System.Drawing.Point(514, 44)
        Me.dgvUnitRight.Name = "dgvUnitRight"
        Me.dgvUnitRight.RowTemplate.Height = 21
        Me.dgvUnitRight.Size = New System.Drawing.Size(273, 502)
        Me.dgvUnitRight.TabIndex = 6
        '
        'rbtnHonkan
        '
        Me.rbtnHonkan.AutoSize = True
        Me.rbtnHonkan.Location = New System.Drawing.Point(327, 15)
        Me.rbtnHonkan.Name = "rbtnHonkan"
        Me.rbtnHonkan.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rbtnHonkan.Size = New System.Drawing.Size(53, 20)
        Me.rbtnHonkan.TabIndex = 7
        Me.rbtnHonkan.Text = "本館"
        Me.rbtnHonkan.UseVisualStyleBackColor = True
        '
        'rbtnAnex
        '
        Me.rbtnAnex.AutoSize = True
        Me.rbtnAnex.Location = New System.Drawing.Point(388, 15)
        Me.rbtnAnex.Name = "rbtnAnex"
        Me.rbtnAnex.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rbtnAnex.Size = New System.Drawing.Size(64, 20)
        Me.rbtnAnex.TabIndex = 8
        Me.rbtnAnex.Text = "ｱﾈｯｸｽ"
        Me.rbtnAnex.UseVisualStyleBackColor = True
        '
        'btnPrevNam
        '
        Me.btnPrevNam.Location = New System.Drawing.Point(459, 8)
        Me.btnPrevNam.Name = "btnPrevNam"
        Me.btnPrevNam.Size = New System.Drawing.Size(67, 29)
        Me.btnPrevNam.TabIndex = 9
        Me.btnPrevNam.Text = "前月氏名"
        Me.btnPrevNam.UseVisualStyleBackColor = True
        '
        'btnRegist
        '
        Me.btnRegist.Location = New System.Drawing.Point(525, 8)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(67, 29)
        Me.btnRegist.TabIndex = 10
        Me.btnRegist.Text = "登　録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        'btnPrevCmpr
        '
        Me.btnPrevCmpr.Location = New System.Drawing.Point(591, 8)
        Me.btnPrevCmpr.Name = "btnPrevCmpr"
        Me.btnPrevCmpr.Size = New System.Drawing.Size(67, 29)
        Me.btnPrevCmpr.TabIndex = 11
        Me.btnPrevCmpr.Text = "前月比"
        Me.btnPrevCmpr.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(723, 8)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(67, 29)
        Me.btnPrint.TabIndex = 12
        Me.btnPrint.Text = "印刷"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(657, 8)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(67, 29)
        Me.btnDelete.TabIndex = 13
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(237, 557)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "測定日："
        '
        'labelUnit1
        '
        Me.labelUnit1.AutoSize = True
        Me.labelUnit1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labelUnit1.Location = New System.Drawing.Point(316, 559)
        Me.labelUnit1.Name = "labelUnit1"
        Me.labelUnit1.Size = New System.Drawing.Size(47, 13)
        Me.labelUnit1.TabIndex = 15
        Me.labelUnit1.Text = "空の家"
        '
        'labelUnit3
        '
        Me.labelUnit3.AutoSize = True
        Me.labelUnit3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labelUnit3.Location = New System.Drawing.Point(316, 613)
        Me.labelUnit3.Name = "labelUnit3"
        Me.labelUnit3.Size = New System.Drawing.Size(47, 13)
        Me.labelUnit3.TabIndex = 19
        Me.labelUnit3.Text = "星の家"
        '
        'labelUnit2
        '
        Me.labelUnit2.AutoSize = True
        Me.labelUnit2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labelUnit2.Location = New System.Drawing.Point(316, 585)
        Me.labelUnit2.Name = "labelUnit2"
        Me.labelUnit2.Size = New System.Drawing.Size(47, 13)
        Me.labelUnit2.TabIndex = 20
        Me.labelUnit2.Text = "森の家"
        '
        'labelUnit5
        '
        Me.labelUnit5.AutoSize = True
        Me.labelUnit5.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labelUnit5.Location = New System.Drawing.Point(578, 583)
        Me.labelUnit5.Name = "labelUnit5"
        Me.labelUnit5.Size = New System.Drawing.Size(47, 13)
        Me.labelUnit5.TabIndex = 23
        Me.labelUnit5.Text = "花の家"
        '
        'labelUnit6
        '
        Me.labelUnit6.AutoSize = True
        Me.labelUnit6.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labelUnit6.Location = New System.Drawing.Point(578, 611)
        Me.labelUnit6.Name = "labelUnit6"
        Me.labelUnit6.Size = New System.Drawing.Size(47, 13)
        Me.labelUnit6.TabIndex = 22
        Me.labelUnit6.Text = "海の家"
        '
        'labelUnit4
        '
        Me.labelUnit4.AutoSize = True
        Me.labelUnit4.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labelUnit4.Location = New System.Drawing.Point(578, 557)
        Me.labelUnit4.Name = "labelUnit4"
        Me.labelUnit4.Size = New System.Drawing.Size(47, 13)
        Me.labelUnit4.TabIndex = 21
        Me.labelUnit4.Text = "月の家"
        '
        'ymdBoxUnit1
        '
        Me.ymdBoxUnit1.DateText = ""
        Me.ymdBoxUnit1.EraText = ""
        Me.ymdBoxUnit1.FirstLabel = "."
        Me.ymdBoxUnit1.FontSize = 9
        Me.ymdBoxUnit1.Location = New System.Drawing.Point(378, 553)
        Me.ymdBoxUnit1.MonthText = ""
        Me.ymdBoxUnit1.Name = "ymdBoxUnit1"
        Me.ymdBoxUnit1.SecondLabel = "."
        Me.ymdBoxUnit1.Size = New System.Drawing.Size(88, 24)
        Me.ymdBoxUnit1.TabIndex = 100
        '
        'ymdBoxUnit3
        '
        Me.ymdBoxUnit3.DateText = ""
        Me.ymdBoxUnit3.EraText = ""
        Me.ymdBoxUnit3.FirstLabel = "."
        Me.ymdBoxUnit3.FontSize = 9
        Me.ymdBoxUnit3.Location = New System.Drawing.Point(378, 609)
        Me.ymdBoxUnit3.MonthText = ""
        Me.ymdBoxUnit3.Name = "ymdBoxUnit3"
        Me.ymdBoxUnit3.SecondLabel = "."
        Me.ymdBoxUnit3.Size = New System.Drawing.Size(88, 24)
        Me.ymdBoxUnit3.TabIndex = 102
        '
        'ymdBoxUnit2
        '
        Me.ymdBoxUnit2.DateText = ""
        Me.ymdBoxUnit2.EraText = ""
        Me.ymdBoxUnit2.FirstLabel = "."
        Me.ymdBoxUnit2.FontSize = 9
        Me.ymdBoxUnit2.Location = New System.Drawing.Point(378, 581)
        Me.ymdBoxUnit2.MonthText = ""
        Me.ymdBoxUnit2.Name = "ymdBoxUnit2"
        Me.ymdBoxUnit2.SecondLabel = "."
        Me.ymdBoxUnit2.Size = New System.Drawing.Size(88, 24)
        Me.ymdBoxUnit2.TabIndex = 101
        '
        'ymdBoxUnit5
        '
        Me.ymdBoxUnit5.DateText = ""
        Me.ymdBoxUnit5.EraText = ""
        Me.ymdBoxUnit5.FirstLabel = "."
        Me.ymdBoxUnit5.FontSize = 9
        Me.ymdBoxUnit5.Location = New System.Drawing.Point(640, 580)
        Me.ymdBoxUnit5.MonthText = ""
        Me.ymdBoxUnit5.Name = "ymdBoxUnit5"
        Me.ymdBoxUnit5.SecondLabel = "."
        Me.ymdBoxUnit5.Size = New System.Drawing.Size(88, 24)
        Me.ymdBoxUnit5.TabIndex = 104
        '
        'ymdBoxUnit6
        '
        Me.ymdBoxUnit6.DateText = ""
        Me.ymdBoxUnit6.EraText = ""
        Me.ymdBoxUnit6.FirstLabel = "."
        Me.ymdBoxUnit6.FontSize = 9
        Me.ymdBoxUnit6.Location = New System.Drawing.Point(640, 608)
        Me.ymdBoxUnit6.MonthText = ""
        Me.ymdBoxUnit6.Name = "ymdBoxUnit6"
        Me.ymdBoxUnit6.SecondLabel = "."
        Me.ymdBoxUnit6.Size = New System.Drawing.Size(88, 24)
        Me.ymdBoxUnit6.TabIndex = 105
        '
        'ymdBoxUnit4
        '
        Me.ymdBoxUnit4.DateText = ""
        Me.ymdBoxUnit4.EraText = ""
        Me.ymdBoxUnit4.FirstLabel = "."
        Me.ymdBoxUnit4.FontSize = 9
        Me.ymdBoxUnit4.Location = New System.Drawing.Point(640, 552)
        Me.ymdBoxUnit4.MonthText = ""
        Me.ymdBoxUnit4.Name = "ymdBoxUnit4"
        Me.ymdBoxUnit4.SecondLabel = "."
        Me.ymdBoxUnit4.Size = New System.Drawing.Size(88, 24)
        Me.ymdBoxUnit4.TabIndex = 103
        '
        '体重管理
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(807, 642)
        Me.Controls.Add(Me.ymdBoxUnit5)
        Me.Controls.Add(Me.ymdBoxUnit6)
        Me.Controls.Add(Me.ymdBoxUnit4)
        Me.Controls.Add(Me.ymdBoxUnit2)
        Me.Controls.Add(Me.ymdBoxUnit3)
        Me.Controls.Add(Me.ymdBoxUnit1)
        Me.Controls.Add(Me.labelUnit5)
        Me.Controls.Add(Me.labelUnit6)
        Me.Controls.Add(Me.labelUnit4)
        Me.Controls.Add(Me.labelUnit2)
        Me.Controls.Add(Me.labelUnit3)
        Me.Controls.Add(Me.labelUnit1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnPrevCmpr)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.btnPrevNam)
        Me.Controls.Add(Me.rbtnAnex)
        Me.Controls.Add(Me.rbtnHonkan)
        Me.Controls.Add(Me.dgvUnitRight)
        Me.Controls.Add(Me.dgvUnitLeft)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.selectUserLabel)
        Me.Controls.Add(Me.dgvUser)
        Me.Name = "体重管理"
        Me.Text = "体重管理"
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvUnitLeft, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvUnitRight, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents selectUserLabel As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dgvUnitLeft As System.Windows.Forms.DataGridView
    Friend WithEvents dgvUnitRight As System.Windows.Forms.DataGridView
    Friend WithEvents rbtnHonkan As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnAnex As System.Windows.Forms.RadioButton
    Friend WithEvents btnPrevNam As System.Windows.Forms.Button
    Friend WithEvents btnRegist As System.Windows.Forms.Button
    Friend WithEvents btnPrevCmpr As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents labelUnit1 As System.Windows.Forms.Label
    Friend WithEvents labelUnit3 As System.Windows.Forms.Label
    Friend WithEvents labelUnit2 As System.Windows.Forms.Label
    Friend WithEvents labelUnit5 As System.Windows.Forms.Label
    Friend WithEvents labelUnit6 As System.Windows.Forms.Label
    Friend WithEvents labelUnit4 As System.Windows.Forms.Label
    Friend WithEvents ymdBoxUnit1 As ymdBox.ymdBox
    Friend WithEvents ymdBoxUnit3 As ymdBox.ymdBox
    Friend WithEvents ymdBoxUnit2 As ymdBox.ymdBox
    Friend WithEvents ymdBoxUnit5 As ymdBox.ymdBox
    Friend WithEvents ymdBoxUnit6 As ymdBox.ymdBox
    Friend WithEvents ymdBoxUnit4 As ymdBox.ymdBox
End Class
