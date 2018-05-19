<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 健康診断
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.lblName = New System.Windows.Forms.Label()
        Me.lblID = New System.Windows.Forms.Label()
        Me.lblHeya = New System.Windows.Forms.Label()
        Me.lblHurigana = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmbKikann = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu1 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu2 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu3 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu4 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu5 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu10 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu9 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu8 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu7 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu6 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu15 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu14 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu13 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu12 = New System.Windows.Forms.ComboBox()
        Me.cmbSaiketu11 = New System.Windows.Forms.ComboBox()
        Me.cmbKennsa1 = New System.Windows.Forms.ComboBox()
        Me.cmbKennsa2 = New System.Windows.Forms.ComboBox()
        Me.cmbKennsa3 = New System.Windows.Forms.ComboBox()
        Me.cmbKennsa4 = New System.Windows.Forms.ComboBox()
        Me.cmbKennsa5 = New System.Windows.Forms.ComboBox()
        Me.cmbKennsa6 = New System.Windows.Forms.ComboBox()
        Me.txtTokki1 = New System.Windows.Forms.TextBox()
        Me.txtTokki2 = New System.Windows.Forms.TextBox()
        Me.txtTokki3 = New System.Windows.Forms.TextBox()
        Me.txtTokki4 = New System.Windows.Forms.TextBox()
        Me.btnTouroku = New System.Windows.Forms.Button()
        Me.btnInnsatu = New System.Windows.Forms.Button()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.btnSakujo = New System.Windows.Forms.Button()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnKuria = New System.Windows.Forms.Button()
        Me.btnKousinn = New System.Windows.Forms.Button()
        Me.DataGridView4 = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.DataGridView5 = New System.Windows.Forms.DataGridView()
        Me.YmdBox1 = New ymdBox.ymdBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(25, 12)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 21
        Me.DataGridView1.Size = New System.Drawing.Size(155, 598)
        Me.DataGridView1.TabIndex = 0
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblName.ForeColor = System.Drawing.Color.Blue
        Me.lblName.Location = New System.Drawing.Point(357, 39)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(47, 19)
        Me.lblName.TabIndex = 6
        Me.lblName.Text = "氏名"
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblID.Location = New System.Drawing.Point(282, 45)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(14, 13)
        Me.lblID.TabIndex = 5
        Me.lblID.Text = "-"
        '
        'lblHeya
        '
        Me.lblHeya.AutoSize = True
        Me.lblHeya.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblHeya.Location = New System.Drawing.Point(258, 16)
        Me.lblHeya.Name = "lblHeya"
        Me.lblHeya.Size = New System.Drawing.Size(59, 13)
        Me.lblHeya.TabIndex = 4
        Me.lblHeya.Text = "部屋番号"
        '
        'lblHurigana
        '
        Me.lblHurigana.AutoSize = True
        Me.lblHurigana.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblHurigana.ForeColor = System.Drawing.Color.Blue
        Me.lblHurigana.Location = New System.Drawing.Point(359, 23)
        Me.lblHurigana.Name = "lblHurigana"
        Me.lblHurigana.Size = New System.Drawing.Size(33, 11)
        Me.lblHurigana.TabIndex = 7
        Me.lblHurigana.Text = "ﾌﾘｶﾞﾅ"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(257, 81)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 16)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "期　　間"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(257, 125)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 16)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "採血項目"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(257, 244)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(62, 16)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "検　　査"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.Location = New System.Drawing.Point(257, 388)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 16)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "特　　記"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.Location = New System.Drawing.Point(257, 498)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(62, 16)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "更　　新"
        '
        'cmbKikann
        '
        Me.cmbKikann.FormattingEnabled = True
        Me.cmbKikann.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbKikann.Items.AddRange(New Object() {"１ヶ月毎", "３ヶ月毎", "６ヶ月毎"})
        Me.cmbKikann.Location = New System.Drawing.Point(347, 79)
        Me.cmbKikann.Name = "cmbKikann"
        Me.cmbKikann.Size = New System.Drawing.Size(134, 20)
        Me.cmbKikann.TabIndex = 13
        '
        'cmbSaiketu1
        '
        Me.cmbSaiketu1.FormattingEnabled = True
        Me.cmbSaiketu1.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu1.Location = New System.Drawing.Point(347, 123)
        Me.cmbSaiketu1.Name = "cmbSaiketu1"
        Me.cmbSaiketu1.Size = New System.Drawing.Size(134, 20)
        Me.cmbSaiketu1.TabIndex = 14
        '
        'cmbSaiketu2
        '
        Me.cmbSaiketu2.FormattingEnabled = True
        Me.cmbSaiketu2.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu2.Location = New System.Drawing.Point(347, 142)
        Me.cmbSaiketu2.Name = "cmbSaiketu2"
        Me.cmbSaiketu2.Size = New System.Drawing.Size(134, 20)
        Me.cmbSaiketu2.TabIndex = 15
        '
        'cmbSaiketu3
        '
        Me.cmbSaiketu3.FormattingEnabled = True
        Me.cmbSaiketu3.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu3.Location = New System.Drawing.Point(347, 161)
        Me.cmbSaiketu3.Name = "cmbSaiketu3"
        Me.cmbSaiketu3.Size = New System.Drawing.Size(134, 20)
        Me.cmbSaiketu3.TabIndex = 16
        '
        'cmbSaiketu4
        '
        Me.cmbSaiketu4.FormattingEnabled = True
        Me.cmbSaiketu4.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu4.Location = New System.Drawing.Point(347, 180)
        Me.cmbSaiketu4.Name = "cmbSaiketu4"
        Me.cmbSaiketu4.Size = New System.Drawing.Size(134, 20)
        Me.cmbSaiketu4.TabIndex = 17
        '
        'cmbSaiketu5
        '
        Me.cmbSaiketu5.FormattingEnabled = True
        Me.cmbSaiketu5.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu5.Location = New System.Drawing.Point(347, 199)
        Me.cmbSaiketu5.Name = "cmbSaiketu5"
        Me.cmbSaiketu5.Size = New System.Drawing.Size(134, 20)
        Me.cmbSaiketu5.TabIndex = 18
        '
        'cmbSaiketu10
        '
        Me.cmbSaiketu10.FormattingEnabled = True
        Me.cmbSaiketu10.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu10.Location = New System.Drawing.Point(480, 199)
        Me.cmbSaiketu10.Name = "cmbSaiketu10"
        Me.cmbSaiketu10.Size = New System.Drawing.Size(133, 20)
        Me.cmbSaiketu10.TabIndex = 23
        '
        'cmbSaiketu9
        '
        Me.cmbSaiketu9.FormattingEnabled = True
        Me.cmbSaiketu9.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu9.Location = New System.Drawing.Point(480, 180)
        Me.cmbSaiketu9.Name = "cmbSaiketu9"
        Me.cmbSaiketu9.Size = New System.Drawing.Size(133, 20)
        Me.cmbSaiketu9.TabIndex = 22
        '
        'cmbSaiketu8
        '
        Me.cmbSaiketu8.FormattingEnabled = True
        Me.cmbSaiketu8.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu8.Location = New System.Drawing.Point(480, 161)
        Me.cmbSaiketu8.Name = "cmbSaiketu8"
        Me.cmbSaiketu8.Size = New System.Drawing.Size(133, 20)
        Me.cmbSaiketu8.TabIndex = 21
        '
        'cmbSaiketu7
        '
        Me.cmbSaiketu7.FormattingEnabled = True
        Me.cmbSaiketu7.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu7.Location = New System.Drawing.Point(480, 142)
        Me.cmbSaiketu7.Name = "cmbSaiketu7"
        Me.cmbSaiketu7.Size = New System.Drawing.Size(133, 20)
        Me.cmbSaiketu7.TabIndex = 20
        '
        'cmbSaiketu6
        '
        Me.cmbSaiketu6.FormattingEnabled = True
        Me.cmbSaiketu6.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu6.Location = New System.Drawing.Point(480, 123)
        Me.cmbSaiketu6.Name = "cmbSaiketu6"
        Me.cmbSaiketu6.Size = New System.Drawing.Size(133, 20)
        Me.cmbSaiketu6.TabIndex = 19
        '
        'cmbSaiketu15
        '
        Me.cmbSaiketu15.FormattingEnabled = True
        Me.cmbSaiketu15.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu15.Location = New System.Drawing.Point(612, 199)
        Me.cmbSaiketu15.Name = "cmbSaiketu15"
        Me.cmbSaiketu15.Size = New System.Drawing.Size(141, 20)
        Me.cmbSaiketu15.TabIndex = 28
        '
        'cmbSaiketu14
        '
        Me.cmbSaiketu14.FormattingEnabled = True
        Me.cmbSaiketu14.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu14.Location = New System.Drawing.Point(612, 180)
        Me.cmbSaiketu14.Name = "cmbSaiketu14"
        Me.cmbSaiketu14.Size = New System.Drawing.Size(141, 20)
        Me.cmbSaiketu14.TabIndex = 27
        '
        'cmbSaiketu13
        '
        Me.cmbSaiketu13.FormattingEnabled = True
        Me.cmbSaiketu13.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu13.Location = New System.Drawing.Point(612, 161)
        Me.cmbSaiketu13.Name = "cmbSaiketu13"
        Me.cmbSaiketu13.Size = New System.Drawing.Size(141, 20)
        Me.cmbSaiketu13.TabIndex = 26
        '
        'cmbSaiketu12
        '
        Me.cmbSaiketu12.FormattingEnabled = True
        Me.cmbSaiketu12.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu12.Location = New System.Drawing.Point(612, 142)
        Me.cmbSaiketu12.Name = "cmbSaiketu12"
        Me.cmbSaiketu12.Size = New System.Drawing.Size(141, 20)
        Me.cmbSaiketu12.TabIndex = 25
        '
        'cmbSaiketu11
        '
        Me.cmbSaiketu11.FormattingEnabled = True
        Me.cmbSaiketu11.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbSaiketu11.Location = New System.Drawing.Point(612, 123)
        Me.cmbSaiketu11.Name = "cmbSaiketu11"
        Me.cmbSaiketu11.Size = New System.Drawing.Size(141, 20)
        Me.cmbSaiketu11.TabIndex = 24
        '
        'cmbKennsa1
        '
        Me.cmbKennsa1.FormattingEnabled = True
        Me.cmbKennsa1.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbKennsa1.Location = New System.Drawing.Point(347, 242)
        Me.cmbKennsa1.Name = "cmbKennsa1"
        Me.cmbKennsa1.Size = New System.Drawing.Size(134, 20)
        Me.cmbKennsa1.TabIndex = 29
        '
        'cmbKennsa2
        '
        Me.cmbKennsa2.FormattingEnabled = True
        Me.cmbKennsa2.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbKennsa2.Location = New System.Drawing.Point(347, 261)
        Me.cmbKennsa2.Name = "cmbKennsa2"
        Me.cmbKennsa2.Size = New System.Drawing.Size(134, 20)
        Me.cmbKennsa2.TabIndex = 30
        '
        'cmbKennsa3
        '
        Me.cmbKennsa3.FormattingEnabled = True
        Me.cmbKennsa3.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbKennsa3.Location = New System.Drawing.Point(347, 280)
        Me.cmbKennsa3.Name = "cmbKennsa3"
        Me.cmbKennsa3.Size = New System.Drawing.Size(134, 20)
        Me.cmbKennsa3.TabIndex = 31
        '
        'cmbKennsa4
        '
        Me.cmbKennsa4.FormattingEnabled = True
        Me.cmbKennsa4.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbKennsa4.Location = New System.Drawing.Point(347, 299)
        Me.cmbKennsa4.Name = "cmbKennsa4"
        Me.cmbKennsa4.Size = New System.Drawing.Size(134, 20)
        Me.cmbKennsa4.TabIndex = 32
        '
        'cmbKennsa5
        '
        Me.cmbKennsa5.FormattingEnabled = True
        Me.cmbKennsa5.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbKennsa5.Location = New System.Drawing.Point(347, 318)
        Me.cmbKennsa5.Name = "cmbKennsa5"
        Me.cmbKennsa5.Size = New System.Drawing.Size(134, 20)
        Me.cmbKennsa5.TabIndex = 33
        '
        'cmbKennsa6
        '
        Me.cmbKennsa6.FormattingEnabled = True
        Me.cmbKennsa6.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.cmbKennsa6.Location = New System.Drawing.Point(347, 337)
        Me.cmbKennsa6.Name = "cmbKennsa6"
        Me.cmbKennsa6.Size = New System.Drawing.Size(134, 20)
        Me.cmbKennsa6.TabIndex = 34
        '
        'txtTokki1
        '
        Me.txtTokki1.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.txtTokki1.Location = New System.Drawing.Point(347, 385)
        Me.txtTokki1.Name = "txtTokki1"
        Me.txtTokki1.Size = New System.Drawing.Size(243, 19)
        Me.txtTokki1.TabIndex = 35
        '
        'txtTokki2
        '
        Me.txtTokki2.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.txtTokki2.Location = New System.Drawing.Point(347, 403)
        Me.txtTokki2.Name = "txtTokki2"
        Me.txtTokki2.Size = New System.Drawing.Size(243, 19)
        Me.txtTokki2.TabIndex = 36
        '
        'txtTokki3
        '
        Me.txtTokki3.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.txtTokki3.Location = New System.Drawing.Point(347, 421)
        Me.txtTokki3.Name = "txtTokki3"
        Me.txtTokki3.Size = New System.Drawing.Size(243, 19)
        Me.txtTokki3.TabIndex = 37
        '
        'txtTokki4
        '
        Me.txtTokki4.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.txtTokki4.Location = New System.Drawing.Point(347, 439)
        Me.txtTokki4.Name = "txtTokki4"
        Me.txtTokki4.Size = New System.Drawing.Size(243, 19)
        Me.txtTokki4.TabIndex = 38
        '
        'btnTouroku
        '
        Me.btnTouroku.Location = New System.Drawing.Point(543, 49)
        Me.btnTouroku.Name = "btnTouroku"
        Me.btnTouroku.Size = New System.Drawing.Size(88, 35)
        Me.btnTouroku.TabIndex = 85
        Me.btnTouroku.Text = "登録"
        Me.btnTouroku.UseVisualStyleBackColor = True
        '
        'btnInnsatu
        '
        Me.btnInnsatu.Location = New System.Drawing.Point(648, 49)
        Me.btnInnsatu.Name = "btnInnsatu"
        Me.btnInnsatu.Size = New System.Drawing.Size(88, 35)
        Me.btnInnsatu.TabIndex = 86
        Me.btnInnsatu.Text = "印刷"
        Me.btnInnsatu.UseVisualStyleBackColor = True
        '
        'DataGridView2
        '
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(668, 295)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowTemplate.Height = 21
        Me.DataGridView2.Size = New System.Drawing.Size(139, 380)
        Me.DataGridView2.TabIndex = 87
        '
        'btnSakujo
        '
        Me.btnSakujo.Location = New System.Drawing.Point(736, 254)
        Me.btnSakujo.Name = "btnSakujo"
        Me.btnSakujo.Size = New System.Drawing.Size(55, 33)
        Me.btnSakujo.TabIndex = 88
        Me.btnSakujo.Text = "削除"
        Me.btnSakujo.UseVisualStyleBackColor = True
        '
        'DataGridView3
        '
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Location = New System.Drawing.Point(701, 65)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.Size = New System.Drawing.Size(10, 10)
        Me.DataGridView3.TabIndex = 89
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(666, 237)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 12)
        Me.Label7.TabIndex = 90
        Me.Label7.Text = "オート保存項目"
        '
        'btnKuria
        '
        Me.btnKuria.Location = New System.Drawing.Point(701, 49)
        Me.btnKuria.Name = "btnKuria"
        Me.btnKuria.Size = New System.Drawing.Size(10, 10)
        Me.btnKuria.TabIndex = 91
        Me.btnKuria.Text = "クリア"
        Me.btnKuria.UseVisualStyleBackColor = True
        '
        'btnKousinn
        '
        Me.btnKousinn.Location = New System.Drawing.Point(717, 49)
        Me.btnKousinn.Name = "btnKousinn"
        Me.btnKousinn.Size = New System.Drawing.Size(10, 10)
        Me.btnKousinn.TabIndex = 103
        Me.btnKousinn.Text = "更新"
        Me.btnKousinn.UseVisualStyleBackColor = True
        '
        'DataGridView4
        '
        Me.DataGridView4.AllowUserToAddRows = False
        Me.DataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView4.Location = New System.Drawing.Point(717, 65)
        Me.DataGridView4.Name = "DataGridView4"
        Me.DataGridView4.RowTemplate.Height = 21
        Me.DataGridView4.Size = New System.Drawing.Size(10, 10)
        Me.DataGridView4.TabIndex = 104
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(258, 46)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(22, 12)
        Me.Label1.TabIndex = 105
        Me.Label1.Text = "ID/"
        '
        'Label8
        '
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Location = New System.Drawing.Point(652, 243)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(10, 1)
        Me.Label8.TabIndex = 106
        Me.Label8.Text = "Label8"
        '
        'Label9
        '
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Location = New System.Drawing.Point(746, 243)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(76, 1)
        Me.Label9.TabIndex = 107
        Me.Label9.Text = "Label9"
        '
        'Label10
        '
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label10.Location = New System.Drawing.Point(652, 684)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(170, 1)
        Me.Label10.TabIndex = 108
        Me.Label10.Text = "Label10"
        '
        'Label11
        '
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label11.Location = New System.Drawing.Point(652, 244)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(1, 440)
        Me.Label11.TabIndex = 109
        Me.Label11.Text = "Label11"
        '
        'Label12
        '
        Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label12.Location = New System.Drawing.Point(821, 244)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(1, 440)
        Me.Label12.TabIndex = 110
        Me.Label12.Text = "Label12"
        '
        'DataGridView5
        '
        Me.DataGridView5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView5.Location = New System.Drawing.Point(621, 65)
        Me.DataGridView5.Name = "DataGridView5"
        Me.DataGridView5.RowTemplate.Height = 21
        Me.DataGridView5.Size = New System.Drawing.Size(10, 10)
        Me.DataGridView5.TabIndex = 111
        '
        'YmdBox1
        '
        Me.YmdBox1.boxType = 3
        Me.YmdBox1.DateText = ""
        Me.YmdBox1.EraText = ""
        Me.YmdBox1.Location = New System.Drawing.Point(347, 482)
        Me.YmdBox1.MonthText = ""
        Me.YmdBox1.Name = "YmdBox1"
        Me.YmdBox1.Size = New System.Drawing.Size(145, 46)
        Me.YmdBox1.TabIndex = 112
        '
        '健康診断
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(851, 718)
        Me.Controls.Add(Me.YmdBox1)
        Me.Controls.Add(Me.btnTouroku)
        Me.Controls.Add(Me.DataGridView5)
        Me.Controls.Add(Me.btnInnsatu)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGridView4)
        Me.Controls.Add(Me.btnKousinn)
        Me.Controls.Add(Me.btnKuria)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.DataGridView3)
        Me.Controls.Add(Me.btnSakujo)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.txtTokki4)
        Me.Controls.Add(Me.txtTokki3)
        Me.Controls.Add(Me.txtTokki2)
        Me.Controls.Add(Me.txtTokki1)
        Me.Controls.Add(Me.cmbKennsa6)
        Me.Controls.Add(Me.cmbKennsa5)
        Me.Controls.Add(Me.cmbKennsa4)
        Me.Controls.Add(Me.cmbKennsa3)
        Me.Controls.Add(Me.cmbKennsa2)
        Me.Controls.Add(Me.cmbKennsa1)
        Me.Controls.Add(Me.cmbSaiketu15)
        Me.Controls.Add(Me.cmbSaiketu14)
        Me.Controls.Add(Me.cmbSaiketu13)
        Me.Controls.Add(Me.cmbSaiketu12)
        Me.Controls.Add(Me.cmbSaiketu11)
        Me.Controls.Add(Me.cmbSaiketu10)
        Me.Controls.Add(Me.cmbSaiketu9)
        Me.Controls.Add(Me.cmbSaiketu8)
        Me.Controls.Add(Me.cmbSaiketu7)
        Me.Controls.Add(Me.cmbSaiketu6)
        Me.Controls.Add(Me.cmbSaiketu5)
        Me.Controls.Add(Me.cmbSaiketu4)
        Me.Controls.Add(Me.cmbSaiketu3)
        Me.Controls.Add(Me.cmbSaiketu2)
        Me.Controls.Add(Me.cmbSaiketu1)
        Me.Controls.Add(Me.cmbKikann)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblHurigana)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.lblID)
        Me.Controls.Add(Me.lblHeya)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "健康診断"
        Me.Text = "健康診断"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents lblHeya As System.Windows.Forms.Label
    Friend WithEvents lblHurigana As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmbKikann As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu2 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu3 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu4 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu5 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu10 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu9 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu8 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu7 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu6 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu15 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu14 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu13 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu12 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSaiketu11 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbKennsa1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbKennsa2 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbKennsa3 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbKennsa4 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbKennsa5 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbKennsa6 As System.Windows.Forms.ComboBox
    Friend WithEvents txtTokki1 As System.Windows.Forms.TextBox
    Friend WithEvents txtTokki2 As System.Windows.Forms.TextBox
    Friend WithEvents txtTokki3 As System.Windows.Forms.TextBox
    Friend WithEvents txtTokki4 As System.Windows.Forms.TextBox
    Friend WithEvents btnTouroku As System.Windows.Forms.Button
    Friend WithEvents btnInnsatu As System.Windows.Forms.Button
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents btnSakujo As System.Windows.Forms.Button
    Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnKuria As System.Windows.Forms.Button
    Friend WithEvents btnKousinn As System.Windows.Forms.Button
    Friend WithEvents DataGridView4 As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents DataGridView5 As System.Windows.Forms.DataGridView
    Friend WithEvents YmdBox1 As ymdBox.ymdBox
End Class
