﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class passwordForm
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
        Me.btnOk = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.passBox = New System.Windows.Forms.TextBox()
        Me.confirmPassBox = New System.Windows.Forms.TextBox()
        Me.newPassBox = New System.Windows.Forms.TextBox()
        Me.errorLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(41, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 12)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "パスワード"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(41, 113)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "変更確認"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(41, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 12)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "新パスワード"
        '
        'btnOk
        '
        Me.btnOk.ForeColor = System.Drawing.Color.Blue
        Me.btnOk.Location = New System.Drawing.Point(93, 179)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 32)
        Me.btnOk.TabIndex = 3
        Me.btnOk.Text = "OK"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.ForeColor = System.Drawing.Color.Blue
        Me.btnCancel.Location = New System.Drawing.Point(177, 179)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 32)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "キャンセル"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'passBox
        '
        Me.passBox.Location = New System.Drawing.Point(121, 31)
        Me.passBox.Name = "passBox"
        Me.passBox.Size = New System.Drawing.Size(130, 19)
        Me.passBox.TabIndex = 0
        '
        'confirmPassBox
        '
        Me.confirmPassBox.Location = New System.Drawing.Point(121, 108)
        Me.confirmPassBox.Name = "confirmPassBox"
        Me.confirmPassBox.Size = New System.Drawing.Size(130, 19)
        Me.confirmPassBox.TabIndex = 2
        '
        'newPassBox
        '
        Me.newPassBox.Location = New System.Drawing.Point(121, 74)
        Me.newPassBox.Name = "newPassBox"
        Me.newPassBox.Size = New System.Drawing.Size(130, 19)
        Me.newPassBox.TabIndex = 1
        '
        'errorLabel
        '
        Me.errorLabel.AutoSize = True
        Me.errorLabel.ForeColor = System.Drawing.Color.Red
        Me.errorLabel.Location = New System.Drawing.Point(41, 145)
        Me.errorLabel.Name = "errorLabel"
        Me.errorLabel.Size = New System.Drawing.Size(38, 12)
        Me.errorLabel.TabIndex = 8
        Me.errorLabel.Text = "Label4"
        Me.errorLabel.Visible = False
        '
        'passwordForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(306, 232)
        Me.Controls.Add(Me.errorLabel)
        Me.Controls.Add(Me.newPassBox)
        Me.Controls.Add(Me.confirmPassBox)
        Me.Controls.Add(Me.passBox)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "passwordForm"
        Me.Text = "管理者パスワード"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents passBox As System.Windows.Forms.TextBox
    Friend WithEvents confirmPassBox As System.Windows.Forms.TextBox
    Friend WithEvents newPassBox As System.Windows.Forms.TextBox
    Friend WithEvents errorLabel As System.Windows.Forms.Label
End Class
