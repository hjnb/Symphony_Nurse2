<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 温度板
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
        Me.lblID = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.YmdBox1 = New ymdBox.ymdBox()
        Me.TimeBox1 = New TimeBox.TimeBox()
        Me.cmbKisaisya = New System.Windows.Forms.ComboBox()
        Me.txtTaionn = New System.Windows.Forms.TextBox()
        Me.txtMyaku = New System.Windows.Forms.TextBox()
        Me.txtKetuatuUe = New System.Windows.Forms.TextBox()
        Me.txtKetuatuSita = New System.Windows.Forms.TextBox()
        Me.txtKettouti = New System.Windows.Forms.TextBox()
        Me.txtSinntyou = New System.Windows.Forms.TextBox()
        Me.txtTaijuu = New System.Windows.Forms.TextBox()
        Me.cmbSyoti = New System.Windows.Forms.ComboBox()
        Me.txtSyoti1 = New System.Windows.Forms.TextBox()
        Me.txtSyoti2 = New System.Windows.Forms.TextBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.btnTouroku = New System.Windows.Forms.Button()
        Me.btnSakujo = New System.Windows.Forms.Button()
        Me.btnInnsatu = New System.Windows.Forms.Button()
        Me.btnDaySita = New System.Windows.Forms.Button()
        Me.btnDayUe = New System.Windows.Forms.Button()
        Me.btnTimeSita = New System.Windows.Forms.Button()
        Me.btnTimeUe = New System.Windows.Forms.Button()
        Me.lstSyoti = New System.Windows.Forms.ListBox()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.btnKousinn = New System.Windows.Forms.Button()
        Me.btnKuria = New System.Windows.Forms.Button()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(43, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "日付"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.Blue
        Me.Label2.Location = New System.Drawing.Point(219, 28)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 12)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "時刻"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(357, 28)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 12)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "記載者"
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblID.ForeColor = System.Drawing.Color.Blue
        Me.lblID.Location = New System.Drawing.Point(543, 25)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(28, 19)
        Me.lblID.TabIndex = 3
        Me.lblID.Text = "ＩＤ"
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblName.ForeColor = System.Drawing.Color.Blue
        Me.lblName.Location = New System.Drawing.Point(633, 25)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(47, 19)
        Me.lblName.TabIndex = 4
        Me.lblName.Text = "氏名"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(43, 74)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(29, 12)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "体温"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(190, 74)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(17, 12)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "脈"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(43, 111)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(53, 12)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "血圧（上）"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(190, 111)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(53, 12)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "血圧（下）"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(43, 148)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(41, 12)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "血糖値"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(190, 148)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(29, 12)
        Me.Label11.TabIndex = 10
        Me.Label11.Text = "身長"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(43, 185)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(29, 12)
        Me.Label12.TabIndex = 11
        Me.Label12.Text = "体重"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(43, 222)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(29, 12)
        Me.Label13.TabIndex = 12
        Me.Label13.Text = "処置"
        '
        'YmdBox1
        '
        Me.YmdBox1.DateText = ""
        Me.YmdBox1.EraText = ""
        Me.YmdBox1.FirstLabel = "."
        Me.YmdBox1.FontSize = 9
        Me.YmdBox1.Location = New System.Drawing.Point(78, 24)
        Me.YmdBox1.MonthText = ""
        Me.YmdBox1.Name = "YmdBox1"
        Me.YmdBox1.SecondLabel = "."
        Me.YmdBox1.Size = New System.Drawing.Size(98, 23)
        Me.YmdBox1.TabIndex = 13
        '
        'TimeBox1
        '
        Me.TimeBox1.HourText = "14"
        Me.TimeBox1.Location = New System.Drawing.Point(254, 23)
        Me.TimeBox1.MinuteText = "36"
        Me.TimeBox1.Name = "TimeBox1"
        Me.TimeBox1.Size = New System.Drawing.Size(67, 23)
        Me.TimeBox1.TabIndex = 14
        '
        'cmbKisaisya
        '
        Me.cmbKisaisya.FormattingEnabled = True
        Me.cmbKisaisya.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbKisaisya.Location = New System.Drawing.Point(404, 25)
        Me.cmbKisaisya.Name = "cmbKisaisya"
        Me.cmbKisaisya.Size = New System.Drawing.Size(88, 20)
        Me.cmbKisaisya.TabIndex = 15
        '
        'txtTaionn
        '
        Me.txtTaionn.Location = New System.Drawing.Point(113, 71)
        Me.txtTaionn.Name = "txtTaionn"
        Me.txtTaionn.Size = New System.Drawing.Size(48, 19)
        Me.txtTaionn.TabIndex = 16
        '
        'txtMyaku
        '
        Me.txtMyaku.Location = New System.Drawing.Point(256, 71)
        Me.txtMyaku.Name = "txtMyaku"
        Me.txtMyaku.Size = New System.Drawing.Size(48, 19)
        Me.txtMyaku.TabIndex = 17
        '
        'txtKetuatuUe
        '
        Me.txtKetuatuUe.Location = New System.Drawing.Point(113, 108)
        Me.txtKetuatuUe.Name = "txtKetuatuUe"
        Me.txtKetuatuUe.Size = New System.Drawing.Size(48, 19)
        Me.txtKetuatuUe.TabIndex = 18
        '
        'txtKetuatuSita
        '
        Me.txtKetuatuSita.Location = New System.Drawing.Point(256, 108)
        Me.txtKetuatuSita.Name = "txtKetuatuSita"
        Me.txtKetuatuSita.Size = New System.Drawing.Size(48, 19)
        Me.txtKetuatuSita.TabIndex = 19
        '
        'txtKettouti
        '
        Me.txtKettouti.Location = New System.Drawing.Point(113, 145)
        Me.txtKettouti.Name = "txtKettouti"
        Me.txtKettouti.Size = New System.Drawing.Size(48, 19)
        Me.txtKettouti.TabIndex = 20
        '
        'txtSinntyou
        '
        Me.txtSinntyou.Location = New System.Drawing.Point(256, 145)
        Me.txtSinntyou.Name = "txtSinntyou"
        Me.txtSinntyou.Size = New System.Drawing.Size(48, 19)
        Me.txtSinntyou.TabIndex = 21
        '
        'txtTaijuu
        '
        Me.txtTaijuu.Location = New System.Drawing.Point(113, 182)
        Me.txtTaijuu.Name = "txtTaijuu"
        Me.txtTaijuu.Size = New System.Drawing.Size(48, 19)
        Me.txtTaijuu.TabIndex = 22
        '
        'cmbSyoti
        '
        Me.cmbSyoti.FormattingEnabled = True
        Me.cmbSyoti.Items.AddRange(New Object() {"治療プラン", "処置", "薬剤内服"})
        Me.cmbSyoti.Location = New System.Drawing.Point(113, 219)
        Me.cmbSyoti.Name = "cmbSyoti"
        Me.cmbSyoti.Size = New System.Drawing.Size(103, 20)
        Me.cmbSyoti.TabIndex = 23
        '
        'txtSyoti1
        '
        Me.txtSyoti1.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.txtSyoti1.Location = New System.Drawing.Point(113, 257)
        Me.txtSyoti1.Name = "txtSyoti1"
        Me.txtSyoti1.Size = New System.Drawing.Size(340, 19)
        Me.txtSyoti1.TabIndex = 24
        '
        'txtSyoti2
        '
        Me.txtSyoti2.Location = New System.Drawing.Point(113, 292)
        Me.txtSyoti2.Name = "txtSyoti2"
        Me.txtSyoti2.ReadOnly = True
        Me.txtSyoti2.Size = New System.Drawing.Size(74, 19)
        Me.txtSyoti2.TabIndex = 25
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(45, 328)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 21
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.DataGridView1.Size = New System.Drawing.Size(756, 340)
        Me.DataGridView1.TabIndex = 26
        '
        'btnTouroku
        '
        Me.btnTouroku.Location = New System.Drawing.Point(618, 71)
        Me.btnTouroku.Name = "btnTouroku"
        Me.btnTouroku.Size = New System.Drawing.Size(82, 45)
        Me.btnTouroku.TabIndex = 27
        Me.btnTouroku.Text = "登録"
        Me.btnTouroku.UseVisualStyleBackColor = True
        '
        'btnSakujo
        '
        Me.btnSakujo.Location = New System.Drawing.Point(618, 132)
        Me.btnSakujo.Name = "btnSakujo"
        Me.btnSakujo.Size = New System.Drawing.Size(82, 45)
        Me.btnSakujo.TabIndex = 28
        Me.btnSakujo.Text = "削除"
        Me.btnSakujo.UseVisualStyleBackColor = True
        '
        'btnInnsatu
        '
        Me.btnInnsatu.Location = New System.Drawing.Point(618, 194)
        Me.btnInnsatu.Name = "btnInnsatu"
        Me.btnInnsatu.Size = New System.Drawing.Size(82, 45)
        Me.btnInnsatu.TabIndex = 29
        Me.btnInnsatu.Text = "印刷"
        Me.btnInnsatu.UseVisualStyleBackColor = True
        '
        'btnDaySita
        '
        Me.btnDaySita.Location = New System.Drawing.Point(176, 34)
        Me.btnDaySita.Name = "btnDaySita"
        Me.btnDaySita.Size = New System.Drawing.Size(24, 22)
        Me.btnDaySita.TabIndex = 84
        Me.btnDaySita.Text = "▼"
        Me.btnDaySita.UseVisualStyleBackColor = True
        '
        'btnDayUe
        '
        Me.btnDayUe.Location = New System.Drawing.Point(176, 13)
        Me.btnDayUe.Name = "btnDayUe"
        Me.btnDayUe.Size = New System.Drawing.Size(24, 22)
        Me.btnDayUe.TabIndex = 83
        Me.btnDayUe.Text = "▲"
        Me.btnDayUe.UseVisualStyleBackColor = True
        '
        'btnTimeSita
        '
        Me.btnTimeSita.Location = New System.Drawing.Point(315, 34)
        Me.btnTimeSita.Name = "btnTimeSita"
        Me.btnTimeSita.Size = New System.Drawing.Size(24, 22)
        Me.btnTimeSita.TabIndex = 86
        Me.btnTimeSita.Text = "▼"
        Me.btnTimeSita.UseVisualStyleBackColor = True
        '
        'btnTimeUe
        '
        Me.btnTimeUe.Location = New System.Drawing.Point(315, 13)
        Me.btnTimeUe.Name = "btnTimeUe"
        Me.btnTimeUe.Size = New System.Drawing.Size(24, 22)
        Me.btnTimeUe.TabIndex = 85
        Me.btnTimeUe.Text = "▲"
        Me.btnTimeUe.UseVisualStyleBackColor = True
        '
        'lstSyoti
        '
        Me.lstSyoti.FormattingEnabled = True
        Me.lstSyoti.ItemHeight = 12
        Me.lstSyoti.Location = New System.Drawing.Point(374, 66)
        Me.lstSyoti.Name = "lstSyoti"
        Me.lstSyoti.Size = New System.Drawing.Size(208, 172)
        Me.lstSyoti.TabIndex = 87
        '
        'DataGridView2
        '
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(869, 34)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowTemplate.Height = 21
        Me.DataGridView2.Size = New System.Drawing.Size(351, 285)
        Me.DataGridView2.TabIndex = 88
        '
        'btnKousinn
        '
        Me.btnKousinn.Location = New System.Drawing.Point(920, 346)
        Me.btnKousinn.Name = "btnKousinn"
        Me.btnKousinn.Size = New System.Drawing.Size(63, 43)
        Me.btnKousinn.TabIndex = 89
        Me.btnKousinn.Text = "更新"
        Me.btnKousinn.UseVisualStyleBackColor = True
        '
        'btnKuria
        '
        Me.btnKuria.Location = New System.Drawing.Point(907, 414)
        Me.btnKuria.Name = "btnKuria"
        Me.btnKuria.Size = New System.Drawing.Size(103, 31)
        Me.btnKuria.TabIndex = 90
        Me.btnKuria.Text = "クリア"
        Me.btnKuria.UseVisualStyleBackColor = True
        '
        'DataGridView3
        '
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Location = New System.Drawing.Point(1095, 415)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.RowTemplate.Height = 21
        Me.DataGridView3.Size = New System.Drawing.Size(167, 214)
        Me.DataGridView3.TabIndex = 91
        '
        '温度板
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1396, 699)
        Me.Controls.Add(Me.DataGridView3)
        Me.Controls.Add(Me.btnKuria)
        Me.Controls.Add(Me.btnKousinn)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.lstSyoti)
        Me.Controls.Add(Me.btnTimeSita)
        Me.Controls.Add(Me.btnTimeUe)
        Me.Controls.Add(Me.btnDaySita)
        Me.Controls.Add(Me.btnDayUe)
        Me.Controls.Add(Me.btnInnsatu)
        Me.Controls.Add(Me.btnSakujo)
        Me.Controls.Add(Me.btnTouroku)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.txtSyoti2)
        Me.Controls.Add(Me.txtSyoti1)
        Me.Controls.Add(Me.cmbSyoti)
        Me.Controls.Add(Me.txtTaijuu)
        Me.Controls.Add(Me.txtSinntyou)
        Me.Controls.Add(Me.txtKettouti)
        Me.Controls.Add(Me.txtKetuatuSita)
        Me.Controls.Add(Me.txtKetuatuUe)
        Me.Controls.Add(Me.txtMyaku)
        Me.Controls.Add(Me.txtTaionn)
        Me.Controls.Add(Me.cmbKisaisya)
        Me.Controls.Add(Me.TimeBox1)
        Me.Controls.Add(Me.YmdBox1)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.lblID)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "温度板"
        Me.Text = "温度板"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents YmdBox1 As ymdBox.ymdBox
    Friend WithEvents TimeBox1 As TimeBox.TimeBox
    Friend WithEvents cmbKisaisya As System.Windows.Forms.ComboBox
    Friend WithEvents txtTaionn As System.Windows.Forms.TextBox
    Friend WithEvents txtMyaku As System.Windows.Forms.TextBox
    Friend WithEvents txtKetuatuUe As System.Windows.Forms.TextBox
    Friend WithEvents txtKetuatuSita As System.Windows.Forms.TextBox
    Friend WithEvents txtKettouti As System.Windows.Forms.TextBox
    Friend WithEvents txtSinntyou As System.Windows.Forms.TextBox
    Friend WithEvents txtTaijuu As System.Windows.Forms.TextBox
    Friend WithEvents cmbSyoti As System.Windows.Forms.ComboBox
    Friend WithEvents txtSyoti1 As System.Windows.Forms.TextBox
    Friend WithEvents txtSyoti2 As System.Windows.Forms.TextBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btnTouroku As System.Windows.Forms.Button
    Friend WithEvents btnSakujo As System.Windows.Forms.Button
    Friend WithEvents btnInnsatu As System.Windows.Forms.Button
    Friend WithEvents btnDaySita As System.Windows.Forms.Button
    Friend WithEvents btnDayUe As System.Windows.Forms.Button
    Friend WithEvents btnTimeSita As System.Windows.Forms.Button
    Friend WithEvents btnTimeUe As System.Windows.Forms.Button
    Friend WithEvents lstSyoti As System.Windows.Forms.ListBox
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents btnKousinn As System.Windows.Forms.Button
    Friend WithEvents btnKuria As System.Windows.Forms.Button
    Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
End Class
