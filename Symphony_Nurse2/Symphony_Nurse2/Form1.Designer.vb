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
        Me.尿バランスToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.体重管理ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.健康診断ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.死亡診断書ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.マスタ登録ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.画面印刷ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ＤＢ整理ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.印刷設定ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.終了ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.利用者選択ToolStripMenuItem, Me.看護日誌ToolStripMenuItem, Me.尿バランスToolStripMenuItem, Me.体重管理ToolStripMenuItem, Me.健康診断ToolStripMenuItem, Me.死亡診断書ToolStripMenuItem, Me.マスタ登録ToolStripMenuItem, Me.画面印刷ToolStripMenuItem, Me.ＤＢ整理ToolStripMenuItem, Me.印刷設定ToolStripMenuItem, Me.終了ToolStripMenuItem})
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
        Me.看護日誌ToolStripMenuItem.Name = "看護日誌ToolStripMenuItem"
        Me.看護日誌ToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.看護日誌ToolStripMenuItem.Text = "看護日誌"
        '
        '尿バランスToolStripMenuItem
        '
        Me.尿バランスToolStripMenuItem.Name = "尿バランスToolStripMenuItem"
        Me.尿バランスToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.尿バランスToolStripMenuItem.Text = "尿バランス"
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
        '死亡診断書ToolStripMenuItem
        '
        Me.死亡診断書ToolStripMenuItem.Name = "死亡診断書ToolStripMenuItem"
        Me.死亡診断書ToolStripMenuItem.Size = New System.Drawing.Size(79, 20)
        Me.死亡診断書ToolStripMenuItem.Text = "死亡診断書"
        '
        'マスタ登録ToolStripMenuItem
        '
        Me.マスタ登録ToolStripMenuItem.Name = "マスタ登録ToolStripMenuItem"
        Me.マスタ登録ToolStripMenuItem.Size = New System.Drawing.Size(70, 20)
        Me.マスタ登録ToolStripMenuItem.Text = "マスタ登録"
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
        'TopForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(817, 524)
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
    Friend WithEvents 尿バランスToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 体重管理ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 健康診断ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 死亡診断書ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents マスタ登録ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 画面印刷ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ＤＢ整理ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 印刷設定ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 終了ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
