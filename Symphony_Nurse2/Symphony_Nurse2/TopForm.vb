Imports System.Reflection

Public Class TopForm

    'とりあえずのデータベースのパス
    Public DB_Nurse2 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\PRIMERGYTX100S1\Hakojun\事務\さかもと\Symphony_Nurse2\Nurse2.mdb"

    '後でちゃんと名前考える
    Private selectUserForm As 利用者選択
    Private aForm As 温度板
    Private bForm As 内服病名
    Private cForm As 体重管理
    Private dForm As 健康診断
    Private eForm As 利用者マスタ
    Private fForm As 処置マスタ
    Private gForm As ユニット居室
    Private hForm As 諸マスタ
    Private iForm As 印刷設定

    ''' <summary>
    ''' コントロールのDoubleBufferedプロパティをTrueにする
    ''' </summary>
    ''' <param name="control">対象のコントロール</param>
    Public Shared Sub EnableDoubleBuffering(control As Control)
        control.GetType().InvokeMember( _
            "DoubleBuffered", _
            BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, _
            Nothing, _
            control, _
            New Object() {True})
    End Sub

    Private Sub TopForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '画面最大化
        Me.WindowState = FormWindowState.Maximized

        '利用者選択フォーム表示
        selectUserForm = New 利用者選択()
        selectUserForm.Owner = Me
        'フォーム内フォームのやり方
        'selectUserForm.TopLevel = False
        'Me.Controls.Add(selectUserForm)
        selectUserForm.Show()

    End Sub

    Private Sub 終了ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 終了ToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub 利用者選択ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 利用者選択ToolStripMenuItem.Click
        If IsNothing(selectUserForm) OrElse selectUserForm.IsDisposed Then
            selectUserForm = New 利用者選択()
            selectUserForm.Owner = Me
            selectUserForm.Show()
        End If
    End Sub

    Private Sub 温度板ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 温度板ToolStripMenuItem.Click
        If IsNothing(aForm) OrElse aForm.IsDisposed Then
            aForm = New 温度板()
            aForm.Owner = Me
            aForm.Show()
        End If
    End Sub

    Private Sub 内服病名ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 内服病名ToolStripMenuItem.Click
        If IsNothing(bForm) OrElse bForm.IsDisposed Then
            bForm = New 内服病名()
            bForm.Owner = Me
            bForm.Show()
        End If
    End Sub

    Private Sub 体重管理ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 体重管理ToolStripMenuItem.Click
        If IsNothing(cForm) OrElse cForm.IsDisposed Then
            cForm = New 体重管理()
            cForm.Owner = Me
            cForm.Show()
        End If
    End Sub

    Private Sub 健康診断ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 健康診断ToolStripMenuItem.Click
        If IsNothing(dForm) OrElse dForm.IsDisposed Then
            dForm = New 健康診断()
            dForm.Owner = Me
            dForm.Show()
        End If
    End Sub

    Private Sub 利用者ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 利用者ToolStripMenuItem.Click
        If IsNothing(eForm) OrElse eForm.IsDisposed Then
            eForm = New 利用者マスタ()
            eForm.Owner = Me
            eForm.Show()
        End If
    End Sub

    Private Sub 処置ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 処置ToolStripMenuItem.Click
        If IsNothing(fForm) OrElse fForm.IsDisposed Then
            fForm = New 処置マスタ()
            fForm.Owner = Me
            fForm.Show()
        End If
    End Sub

    Private Sub ユニット居室ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ユニット居室ToolStripMenuItem.Click
        If IsNothing(gForm) OrElse gForm.IsDisposed Then
            gForm = New ユニット居室()
            gForm.Owner = Me
            gForm.Show()
        End If
    End Sub

    Private Sub 諸ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 諸ToolStripMenuItem.Click
        If IsNothing(hForm) OrElse hForm.IsDisposed Then
            hForm = New 諸マスタ()
            hForm.Owner = Me
            hForm.Show()
        End If
    End Sub

    Private Sub 画面印刷ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 画面印刷ToolStripMenuItem.Click
        
    End Sub

    Private Sub ＤＢ整理ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ＤＢ整理ToolStripMenuItem.Click

    End Sub

    Private Sub 印刷設定ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 印刷設定ToolStripMenuItem.Click
        If IsNothing(iForm) OrElse iForm.IsDisposed Then
            iForm = New 印刷設定()
            iForm.Owner = Me
            iForm.Show()
        End If
    End Sub

End Class
