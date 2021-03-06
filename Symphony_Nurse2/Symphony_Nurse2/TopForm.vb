﻿Imports System.Reflection

Public Class TopForm

    'データベースのパス
    Public dbFilePath As String = My.Application.Info.DirectoryPath & "\Nurse2.mdb"
    Public DB_Nurse2 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFilePath

    'エクセルのパス
    Public excelFilePass As String = My.Application.Info.DirectoryPath & "\Nurse2.xls"

    '.iniファイルのパス
    Public iniFilePath As String = My.Application.Info.DirectoryPath & "\Nurse2.ini"

    '後でちゃんと名前考える
    Public selectUserForm As 利用者選択
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
        'データベース、エクセル、構成ファイルの存在チェック
        If Not System.IO.File.Exists(dbFilePath) Then
            MsgBox("データベースファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        If Not System.IO.File.Exists(excelFilePass) Then
            MsgBox("エクセルファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        If Not System.IO.File.Exists(iniFilePath) Then
            MsgBox("構成ファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        '画面最大化
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub TopForm_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        '利用者選択フォーム表示
        selectUserForm = New 利用者選択()
        selectUserForm.Owner = Me
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
        PrintForm()
    End Sub

    Private Sub 印刷設定ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 印刷設定ToolStripMenuItem.Click
        If IsNothing(iForm) OrElse iForm.IsDisposed Then
            iForm = New 印刷設定()
            iForm.Owner = Me
            iForm.Show()
        End If
    End Sub

    Private memoryImage As Bitmap

    Public Function getCaptureImage() As Bitmap
        'Bitmapの作成
        Dim bmp As New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
        'Graphicsの作成
        Dim g As Graphics = Graphics.FromImage(bmp)
        g.CopyFromScreen(New Point(Me.Left + 8, Me.Top), New Point(0, 0), New Size(Me.Width - 16, Me.Height - 16))
        '解放
        g.Dispose()

        Return bmp
    End Function

    Public Sub PrintForm()
        'フォームのイメージを取得する
        memoryImage = getCaptureImage()
        'フォームのイメージを印刷する
        Dim PrintDocument As New System.Drawing.Printing.PrintDocument
        PrintDocument.DefaultPageSettings.Landscape = True
        AddHandler PrintDocument.PrintPage, AddressOf PrintDocument_PrintPage
        PrintDocument.Print()
        memoryImage.Dispose()
    End Sub

    Private Sub PrintDocument_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        '拡大率を指定
        Dim zoom As Double = 1
        Dim padding As Double = 20
        If memoryImage.Width > e.Graphics.VisibleClipBounds.Width Then
            zoom = e.Graphics.VisibleClipBounds.Width / memoryImage.Width
        End If

        If (memoryImage.Height + padding) * zoom > e.Graphics.VisibleClipBounds.Height Then
            zoom = e.Graphics.VisibleClipBounds.Height / (memoryImage.Height + padding)
        End If

        e.Graphics.DrawImage(memoryImage, 0, CInt(padding), CInt(memoryImage.Width * zoom), CInt(memoryImage.Height * zoom))
    End Sub
End Class
