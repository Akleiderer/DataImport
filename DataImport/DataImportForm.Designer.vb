<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DataImportForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DataImportForm))
        Me.MinLabel = New System.Windows.Forms.Label()
        Me.CloseLabel = New System.Windows.Forms.Label()
        Me.Title = New System.Windows.Forms.Label()
        Me.FolderInputBox = New System.Windows.Forms.TextBox()
        Me.FileList = New System.Windows.Forms.ListBox()
        Me.SelectedList = New System.Windows.Forms.ListBox()
        Me.TextOutput = New System.Windows.Forms.TextBox()
        Me.AddSingle = New System.Windows.Forms.Label()
        Me.AddAll = New System.Windows.Forms.Label()
        Me.RemoveSingle = New System.Windows.Forms.Label()
        Me.RemoveAll = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.InputBrowseButton = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.OutputBrowseButton = New System.Windows.Forms.Button()
        Me.FolderOutputBox = New System.Windows.Forms.TextBox()
        Me.ConvertButton = New System.Windows.Forms.Label()
        Me.InputFolderDialogue = New System.Windows.Forms.FolderBrowserDialog()
        Me.OutputFolderDialogue = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MinLabel
        '
        Me.MinLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MinLabel.AutoSize = True
        Me.MinLabel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.MinLabel.Font = New System.Drawing.Font("Consolas", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MinLabel.ForeColor = System.Drawing.Color.White
        Me.MinLabel.Location = New System.Drawing.Point(519, 0)
        Me.MinLabel.Name = "MinLabel"
        Me.MinLabel.Size = New System.Drawing.Size(37, 41)
        Me.MinLabel.TabIndex = 0
        Me.MinLabel.Text = "-"
        Me.MinLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CloseLabel
        '
        Me.CloseLabel.AutoSize = True
        Me.CloseLabel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CloseLabel.Dock = System.Windows.Forms.DockStyle.Right
        Me.CloseLabel.Font = New System.Drawing.Font("Consolas", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CloseLabel.ForeColor = System.Drawing.Color.White
        Me.CloseLabel.Location = New System.Drawing.Point(563, 0)
        Me.CloseLabel.Name = "CloseLabel"
        Me.CloseLabel.Size = New System.Drawing.Size(37, 41)
        Me.CloseLabel.TabIndex = 1
        Me.CloseLabel.Text = "x"
        Me.CloseLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Title
        '
        Me.Title.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.Title.Font = New System.Drawing.Font("Arial", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Title.ForeColor = System.Drawing.Color.White
        Me.Title.Location = New System.Drawing.Point(175, 10)
        Me.Title.Name = "Title"
        Me.Title.Size = New System.Drawing.Size(250, 50)
        Me.Title.TabIndex = 0
        Me.Title.Text = "Data Import"
        Me.Title.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'FolderInputBox
        '
        Me.FolderInputBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.749999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FolderInputBox.Location = New System.Drawing.Point(50, 125)
        Me.FolderInputBox.Name = "FolderInputBox"
        Me.FolderInputBox.Size = New System.Drawing.Size(390, 22)
        Me.FolderInputBox.TabIndex = 1
        '
        'FileList
        '
        Me.FileList.FormattingEnabled = True
        Me.FileList.Location = New System.Drawing.Point(50, 180)
        Me.FileList.Name = "FileList"
        Me.FileList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.FileList.Size = New System.Drawing.Size(225, 238)
        Me.FileList.Sorted = True
        Me.FileList.TabIndex = 4
        Me.FileList.TabStop = False
        '
        'SelectedList
        '
        Me.SelectedList.FormattingEnabled = True
        Me.SelectedList.Location = New System.Drawing.Point(325, 180)
        Me.SelectedList.Name = "SelectedList"
        Me.SelectedList.Size = New System.Drawing.Size(225, 238)
        Me.SelectedList.TabIndex = 5
        Me.SelectedList.TabStop = False
        '
        'TextOutput
        '
        Me.TextOutput.AcceptsReturn = True
        Me.TextOutput.AcceptsTab = True
        Me.TextOutput.BackColor = System.Drawing.Color.Silver
        Me.TextOutput.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextOutput.Location = New System.Drawing.Point(50, 542)
        Me.TextOutput.MaxLength = 0
        Me.TextOutput.Multiline = True
        Me.TextOutput.Name = "TextOutput"
        Me.TextOutput.ReadOnly = True
        Me.TextOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextOutput.Size = New System.Drawing.Size(500, 135)
        Me.TextOutput.TabIndex = 7
        '
        'AddSingle
        '
        Me.AddSingle.BackColor = System.Drawing.Color.Gray
        Me.AddSingle.Font = New System.Drawing.Font("Cooper Black", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddSingle.ForeColor = System.Drawing.Color.Blue
        Me.AddSingle.Location = New System.Drawing.Point(280, 225)
        Me.AddSingle.Margin = New System.Windows.Forms.Padding(0)
        Me.AddSingle.Name = "AddSingle"
        Me.AddSingle.Size = New System.Drawing.Size(40, 25)
        Me.AddSingle.TabIndex = 8
        Me.AddSingle.Text = ">"
        Me.AddSingle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'AddAll
        '
        Me.AddAll.BackColor = System.Drawing.Color.Gray
        Me.AddAll.Font = New System.Drawing.Font("Cooper Black", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddAll.ForeColor = System.Drawing.Color.Blue
        Me.AddAll.Location = New System.Drawing.Point(280, 255)
        Me.AddAll.Margin = New System.Windows.Forms.Padding(0)
        Me.AddAll.Name = "AddAll"
        Me.AddAll.Size = New System.Drawing.Size(40, 25)
        Me.AddAll.TabIndex = 9
        Me.AddAll.Text = ">>"
        Me.AddAll.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RemoveSingle
        '
        Me.RemoveSingle.BackColor = System.Drawing.Color.Gray
        Me.RemoveSingle.Font = New System.Drawing.Font("Cooper Black", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RemoveSingle.ForeColor = System.Drawing.Color.Blue
        Me.RemoveSingle.Location = New System.Drawing.Point(280, 360)
        Me.RemoveSingle.Margin = New System.Windows.Forms.Padding(0)
        Me.RemoveSingle.Name = "RemoveSingle"
        Me.RemoveSingle.Size = New System.Drawing.Size(40, 25)
        Me.RemoveSingle.TabIndex = 11
        Me.RemoveSingle.Text = "<"
        Me.RemoveSingle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.RemoveSingle.UseMnemonic = False
        '
        'RemoveAll
        '
        Me.RemoveAll.BackColor = System.Drawing.Color.Gray
        Me.RemoveAll.Font = New System.Drawing.Font("Cooper Black", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RemoveAll.ForeColor = System.Drawing.Color.Blue
        Me.RemoveAll.Location = New System.Drawing.Point(280, 330)
        Me.RemoveAll.Margin = New System.Windows.Forms.Padding(0)
        Me.RemoveAll.Name = "RemoveAll"
        Me.RemoveAll.Size = New System.Drawing.Size(40, 25)
        Me.RemoveAll.TabIndex = 10
        Me.RemoveAll.Text = "<<"
        Me.RemoveAll.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(325, 158)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(225, 19)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Selected Files"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'InputBrowseButton
        '
        Me.InputBrowseButton.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InputBrowseButton.Location = New System.Drawing.Point(450, 125)
        Me.InputBrowseButton.Margin = New System.Windows.Forms.Padding(0)
        Me.InputBrowseButton.Name = "InputBrowseButton"
        Me.InputBrowseButton.Size = New System.Drawing.Size(100, 25)
        Me.InputBrowseButton.TabIndex = 2
        Me.InputBrowseButton.Text = "Browse"
        Me.InputBrowseButton.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(50, 90)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(100, 25)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "Input Folder"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(50, 435)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(69, 25)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Output Folder"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'OutputBrowseButton
        '
        Me.OutputBrowseButton.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OutputBrowseButton.Location = New System.Drawing.Point(450, 458)
        Me.OutputBrowseButton.Margin = New System.Windows.Forms.Padding(0)
        Me.OutputBrowseButton.Name = "OutputBrowseButton"
        Me.OutputBrowseButton.Size = New System.Drawing.Size(100, 23)
        Me.OutputBrowseButton.TabIndex = 4
        Me.OutputBrowseButton.Text = "Browse"
        Me.OutputBrowseButton.UseVisualStyleBackColor = True
        '
        'FolderOutputBox
        '
        Me.FolderOutputBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FolderOutputBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.749999!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FolderOutputBox.Location = New System.Drawing.Point(54, 458)
        Me.FolderOutputBox.Name = "FolderOutputBox"
        Me.FolderOutputBox.Size = New System.Drawing.Size(390, 22)
        Me.FolderOutputBox.TabIndex = 3
        '
        'ConvertButton
        '
        Me.ConvertButton.BackColor = System.Drawing.Color.Gray
        Me.ConvertButton.Font = New System.Drawing.Font("Tahoma", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ConvertButton.ForeColor = System.Drawing.Color.Blue
        Me.ConvertButton.Location = New System.Drawing.Point(200, 500)
        Me.ConvertButton.Name = "ConvertButton"
        Me.ConvertButton.Size = New System.Drawing.Size(200, 35)
        Me.ConvertButton.TabIndex = 21
        Me.ConvertButton.Text = "Convert Data"
        Me.ConvertButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'InputFolderDialogue
        '
        Me.InputFolderDialogue.RootFolder = System.Environment.SpecialFolder.MyComputer
        '
        'OutputFolderDialogue
        '
        Me.OutputFolderDialogue.RootFolder = System.Environment.SpecialFolder.MyComputer
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(50, 158)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(225, 19)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Valid Files Found"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox1.ErrorImage = Nothing
        Me.PictureBox1.InitialImage = Nothing
        Me.PictureBox1.Location = New System.Drawing.Point(50, 10)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(64, 64)
        Me.PictureBox1.TabIndex = 22
        Me.PictureBox1.TabStop = False
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.Silver
        Me.TextBox1.Location = New System.Drawing.Point(548, 683)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(52, 20)
        Me.TextBox1.TabIndex = 23
        Me.TextBox1.Text = "v. 1.0.5"
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'DataImportForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(50, Byte), Integer), CType(CType(50, Byte), Integer), CType(CType(50, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(600, 700)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.ConvertButton)
        Me.Controls.Add(Me.FolderOutputBox)
        Me.Controls.Add(Me.OutputBrowseButton)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.InputBrowseButton)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.RemoveSingle)
        Me.Controls.Add(Me.RemoveAll)
        Me.Controls.Add(Me.AddAll)
        Me.Controls.Add(Me.AddSingle)
        Me.Controls.Add(Me.TextOutput)
        Me.Controls.Add(Me.SelectedList)
        Me.Controls.Add(Me.FileList)
        Me.Controls.Add(Me.FolderInputBox)
        Me.Controls.Add(Me.Title)
        Me.Controls.Add(Me.CloseLabel)
        Me.Controls.Add(Me.MinLabel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "DataImportForm"
        Me.Text = "Data Import"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MinLabel As Label
    Friend WithEvents CloseLabel As Label
    Friend WithEvents Title As Label
    Friend WithEvents FolderInputBox As TextBox
    Friend WithEvents FileList As ListBox
    Friend WithEvents SelectedList As ListBox
    Friend WithEvents TextOutput As TextBox
    Friend WithEvents AddSingle As Label
    Friend WithEvents AddAll As Label
    Friend WithEvents RemoveSingle As Label
    Friend WithEvents RemoveAll As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents InputBrowseButton As Button
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents OutputBrowseButton As Button
    Friend WithEvents FolderOutputBox As TextBox
    Friend WithEvents ConvertButton As Label
    Friend WithEvents InputFolderDialogue As FolderBrowserDialog
    Friend WithEvents OutputFolderDialogue As FolderBrowserDialog
    Friend WithEvents Label5 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents TextBox1 As TextBox
    Public WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
