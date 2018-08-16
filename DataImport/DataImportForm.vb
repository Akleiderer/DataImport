Public Class DataImportForm

    Dim drag As Boolean
    Dim mousex As Integer
    Dim mousey As Integer
    Dim inputdir As IO.DirectoryInfo
    Dim validexts() As String

    Private Sub DataImportForm_MouseMove(sender As Object, e As EventArgs) Handles MyBase.MouseMove
        ' If drag is true, then move the form
        If drag Then
            Me.Top = Windows.Forms.Cursor.Position.Y - mousey
            Me.Left = Windows.Forms.Cursor.Position.X - mousex
        End If

    End Sub
    Private Sub DataImportForm_MouseDown(sender As Object, e As EventArgs) Handles MyBase.MouseDown
        ' Sets drag to true and stores mouse position
        drag = True
        mousex = Windows.Forms.Cursor.Position.X - Me.Left
        mousey = Windows.Forms.Cursor.Position.Y - Me.Top

    End Sub
    Private Sub DataImportForm_MouseUp(sender As Object, e As EventArgs) Handles MyBase.MouseUp
        ' Sets drag to false
        drag = False

    End Sub

    Private Sub InputBrowseButton_Click(sender As Object, e As EventArgs) Handles InputBrowseButton.Click
        If (InputFolderDialogue.ShowDialog() = DialogResult.OK) Then
            FolderInputBox.Text = InputFolderDialogue.SelectedPath
            FolderOutputBox.Text = String.Concat(InputFolderDialogue.SelectedPath, "\Xlsx Files\")
        End If
    End Sub

    Private Sub OutputBrowseButton_Click(sender As Object, e As EventArgs) Handles OutputBrowseButton.Click
        If (OutputFolderDialogue.ShowDialog() = DialogResult.OK) Then
            FolderOutputBox.Text = OutputFolderDialogue.SelectedPath
        End If
    End Sub

    Private Sub CloseLabel_Click(sender As Object, e As EventArgs) Handles CloseLabel.Click
        Me.Close()
    End Sub

    Private Sub MinLabel_Click(sender As Object, e As EventArgs) Handles MinLabel.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub FolderInputBox_TextChanged(sender As Object, e As EventArgs) Handles FolderInputBox.TextChanged
        inputdir = New IO.DirectoryInfo(FolderInputBox.Text)
        If inputdir.Exists Then
            PopulateFileList(inputdir)
        End If
    End Sub

    Private Sub PopulateFileList(dir As IO.DirectoryInfo)
        FileList.Items.Clear()
        For Each file In dir.EnumerateFiles()
            If validexts.Contains(file.Extension()) Then
                FileList.Items.Add(file.Name())
            End If
        Next
    End Sub

    Private Sub DataImportForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        validexts = {".res", ".mdb"}
    End Sub
End Class
