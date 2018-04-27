<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TopForm
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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.利用者選択ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.看護日誌ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.温度板ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.内服病名ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.体重管理ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.健康診断ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.マスタ登録ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.利用者ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.処置ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ユニット居室ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.諸ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.画面印刷ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ＤＢ整理ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.印刷設定ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.終了ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.userId = New System.Windows.Forms.TextBox()
        Me.userNam = New System.Windows.Forms.TextBox()
        Me.userDsp = New System.Windows.Forms.TextBox()
        Me.userKana = New System.Windows.Forms.TextBox()
        Me.userBirth = New System.Windows.Forms.TextBox()
        Me.userSex = New System.Windows.Forms.TextBox()
        Me.userKaigo = New System.Windows.Forms.TextBox()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.利用者選択ToolStripMenuItem, Me.看護日誌ToolStripMenuItem, Me.体重管理ToolStripMenuItem, Me.健康診断ToolStripMenuItem, Me.マスタ登録ToolStripMenuItem, Me.画面印刷ToolStripMenuItem, Me.ＤＢ整理ToolStripMenuItem, Me.印刷設定ToolStripMenuItem, Me.終了ToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(817, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        '利用者選択ToolStripMenuItem
        '
        Me.利用者選択ToolStripMenuItem.Name = "利用者選択ToolStripMenuItem"
        Me.利用者選択ToolStripMenuItem.Size = New System.Drawing.Size(79, 20)
        Me.利用者選択ToolStripMenuItem.Text = "利用者選択"
        '
        '看護日誌ToolStripMenuItem
        '
        Me.看護日誌ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.温度板ToolStripMenuItem, Me.内服病名ToolStripMenuItem})
        Me.看護日誌ToolStripMenuItem.Name = "看護日誌ToolStripMenuItem"
        Me.看護日誌ToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.看護日誌ToolStripMenuItem.Text = "看護日誌"
        '
        '温度板ToolStripMenuItem
        '
        Me.温度板ToolStripMenuItem.Name = "温度板ToolStripMenuItem"
        Me.温度板ToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.温度板ToolStripMenuItem.Text = "温度板"
        '
        '内服病名ToolStripMenuItem
        '
        Me.内服病名ToolStripMenuItem.Name = "内服病名ToolStripMenuItem"
        Me.内服病名ToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.内服病名ToolStripMenuItem.Text = "内服病名"
        '
        '体重管理ToolStripMenuItem
        '
        Me.体重管理ToolStripMenuItem.Name = "体重管理ToolStripMenuItem"
        Me.体重管理ToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.体重管理ToolStripMenuItem.Text = "体重管理"
        '
        '健康診断ToolStripMenuItem
        '
        Me.健康診断ToolStripMenuItem.Name = "健康診断ToolStripMenuItem"
        Me.健康診断ToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.健康診断ToolStripMenuItem.Text = "健康診断"
        '
        'マスタ登録ToolStripMenuItem
        '
        Me.マスタ登録ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.利用者ToolStripMenuItem, Me.処置ToolStripMenuItem, Me.ユニット居室ToolStripMenuItem, Me.諸ToolStripMenuItem})
        Me.マスタ登録ToolStripMenuItem.Name = "マスタ登録ToolStripMenuItem"
        Me.マスタ登録ToolStripMenuItem.Size = New System.Drawing.Size(70, 20)
        Me.マスタ登録ToolStripMenuItem.Text = "マスタ登録"
        '
        '利用者ToolStripMenuItem
        '
        Me.利用者ToolStripMenuItem.Name = "利用者ToolStripMenuItem"
        Me.利用者ToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.利用者ToolStripMenuItem.Text = "利用者"
        '
        '処置ToolStripMenuItem
        '
        Me.処置ToolStripMenuItem.Name = "処置ToolStripMenuItem"
        Me.処置ToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.処置ToolStripMenuItem.Text = "処置"
        '
        'ユニット居室ToolStripMenuItem
        '
        Me.ユニット居室ToolStripMenuItem.Name = "ユニット居室ToolStripMenuItem"
        Me.ユニット居室ToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.ユニット居室ToolStripMenuItem.Text = "ユニット居室"
        '
        '諸ToolStripMenuItem
        '
        Me.諸ToolStripMenuItem.Name = "諸ToolStripMenuItem"
        Me.諸ToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.諸ToolStripMenuItem.Text = "諸"
        '
        '画面印刷ToolStripMenuItem
        '
        Me.画面印刷ToolStripMenuItem.Name = "画面印刷ToolStripMenuItem"
        Me.画面印刷ToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.画面印刷ToolStripMenuItem.Text = "画面印刷"
        '
        'ＤＢ整理ToolStripMenuItem
        '
        Me.ＤＢ整理ToolStripMenuItem.Name = "ＤＢ整理ToolStripMenuItem"
        Me.ＤＢ整理ToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.ＤＢ整理ToolStripMenuItem.Text = "ＤＢ整理"
        '
        '印刷設定ToolStripMenuItem
        '
        Me.印刷設定ToolStripMenuItem.Name = "印刷設定ToolStripMenuItem"
        Me.印刷設定ToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.印刷設定ToolStripMenuItem.Text = "印刷設定"
        '
        '終了ToolStripMenuItem
        '
        Me.終了ToolStripMenuItem.Name = "終了ToolStripMenuItem"
        Me.終了ToolStripMenuItem.Size = New System.Drawing.Size(43, 20)
        Me.終了ToolStripMenuItem.Text = "終了"
        '
        'userId
        '
        Me.userId.Location = New System.Drawing.Point(624, 225)
        Me.userId.Name = "userId"
        Me.userId.Size = New System.Drawing.Size(100, 19)
        Me.userId.TabIndex = 1
        '
        'userNam
        '
        Me.userNam.Location = New System.Drawing.Point(624, 259)
        Me.userNam.Name = "userNam"
        Me.userNam.Size = New System.Drawing.Size(100, 19)
        Me.userNam.TabIndex = 2
        '
        'userDsp
        '
        Me.userDsp.Location = New System.Drawing.Point(624, 327)
        Me.userDsp.Name = "userDsp"
        Me.userDsp.Size = New System.Drawing.Size(100, 19)
        Me.userDsp.TabIndex = 4
        '
        'userKana
        '
        Me.userKana.Location = New System.Drawing.Point(624, 293)
        Me.userKana.Name = "userKana"
        Me.userKana.Size = New System.Drawing.Size(100, 19)
        Me.userKana.TabIndex = 3
        '
        'userBirth
        '
        Me.userBirth.Location = New System.Drawing.Point(624, 398)
        Me.userBirth.Name = "userBirth"
        Me.userBirth.Size = New System.Drawing.Size(100, 19)
        Me.userBirth.TabIndex = 6
        '
        'userSex
        '
        Me.userSex.Location = New System.Drawing.Point(624, 363)
        Me.userSex.Name = "userSex"
        Me.userSex.Size = New System.Drawing.Size(100, 19)
        Me.userSex.TabIndex = 5
        '
        'userKaigo
        '
        Me.userKaigo.Location = New System.Drawing.Point(624, 433)
        Me.userKaigo.Name = "userKaigo"
        Me.userKaigo.Size = New System.Drawing.Size(100, 19)
        Me.userKaigo.TabIndex = 7
        '
        'TopForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(171, Byte), Integer), CType(CType(171, Byte), Integer), CType(CType(171, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(817, 524)
        Me.Controls.Add(Me.userKaigo)
        Me.Controls.Add(Me.userBirth)
        Me.Controls.Add(Me.userSex)
        Me.Controls.Add(Me.userDsp)
        Me.Controls.Add(Me.userKana)
        Me.Controls.Add(Me.userNam)
        Me.Controls.Add(Me.userId)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "TopForm"
        Me.Text = "Nurse2"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents 利用者選択ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 看護日誌ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 体重管理ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 健康診断ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents マスタ登録ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 画面印刷ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ＤＢ整理ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 印刷設定ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 終了ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 温度板ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 内服病名ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 利用者ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 処置ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ユニット居室ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 諸ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents userId As System.Windows.Forms.TextBox
    Friend WithEvents userNam As System.Windows.Forms.TextBox
    Friend WithEvents userDsp As System.Windows.Forms.TextBox
    Friend WithEvents userKana As System.Windows.Forms.TextBox
    Friend WithEvents userBirth As System.Windows.Forms.TextBox
    Friend WithEvents userSex As System.Windows.Forms.TextBox
    Friend WithEvents userKaigo As System.Windows.Forms.TextBox

End Class
