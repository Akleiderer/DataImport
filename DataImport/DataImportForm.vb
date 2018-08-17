
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel

Public Class DataImportForm

    Dim drag As Boolean
    Dim mousex As Integer
    Dim mousey As Integer
    Dim inputdir As IO.DirectoryInfo
    Dim validexts() As String
    Dim tempfiles As New Dictionary(Of String, IO.FileInfo)
    Dim arbinfiles As New List(Of IO.FileInfo)

    Private Sub DataImportForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        validexts = {".res", ".mdb"}
    End Sub

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
        If IO.Directory.Exists(FolderInputBox.Text) Then
            inputdir = New IO.DirectoryInfo(FolderInputBox.Text)
            PopulateFileList(inputdir)
        End If
    End Sub

    Private Sub PopulateFileList(dir As IO.DirectoryInfo)
        FileList.Items.Clear()
        For Each file In dir.EnumerateFiles()
            If validexts.Contains(file.Extension()) Then
                tempfiles.Add(file.Name, file)
                FileList.Items.Add(file.Name())
            End If
        Next
    End Sub

    Private Sub AddSingle_Click(sender As Object, e As EventArgs) Handles AddSingle.Click
        For Each item In FileList.SelectedItems
            If Not SelectedList.Items.Contains(item) Then
                SelectedList.Items.Add(item)
            End If
        Next
    End Sub

    Private Sub AddAll_Click(sender As Object, e As EventArgs) Handles AddAll.Click
        For Each item In FileList.Items
            If Not SelectedList.Items.Contains(item) Then
                SelectedList.Items.Add(item)
            End If
        Next
    End Sub

    Private Sub RemoveAll_Click(sender As Object, e As EventArgs) Handles RemoveAll.Click
        SelectedList.Items.Clear()
    End Sub

    Private Sub RemoveSingle_Click(sender As Object, e As EventArgs) Handles RemoveSingle.Click
        Dim items
        items = SelectedList.SelectedItems()
        For Each item In items
            SelectedList.Items.Remove(item)
        Next
    End Sub

    Private Sub ConvertButton_Click(sender As Object, e As EventArgs) Handles ConvertButton.Click
        If Not IO.Directory.Exists(FolderOutputBox.Text) Then
            Try
                IO.Directory.CreateDirectory(FolderOutputBox.Text)
            Catch ex As Exception
                MsgBox("Please select a valid output directory" & vbCrLf & " Error: " & ex.Message)
            End Try
        End If

        For Each filename In SelectedList.Items
            If {".res", ".mdb"}.Contains(tempfiles(filename).Extension) Then
                arbinfiles.Add(tempfiles(filename))
            End If
        Next
        TextOutput.AppendText(String.Concat("Converting files:", vbCrLf))
        For Each file In arbinfiles
            TextOutput.AppendText(String.Concat(file.Name, vbCrLf))
        Next

        Arbin.Convert(arbinfiles, FolderOutputBox.Text)
    End Sub


End Class

Module Arbin
    Public Sub Convert(ByVal files As List(Of IO.FileInfo), outputpath As String)
        Dim normaltable As New DataTable("Normal Table")
        Dim statstable As New DataTable("Statistics Table")
        Dim outputdir As New IO.DirectoryInfo(outputpath)


        For Each file In files
            DataImportForm.TextOutput.AppendText("Pulling tables for " & file.Name & "." & vbCrLf)

            normaltable = GetTable(file, normaltable.TableName())
            statstable = GetTable(file, statstable.TableName())

            DataImportForm.TextOutput.AppendText("Exporting tables to Excel." & vbCrLf)
            If Export.ToExcel({normaltable, statstable}, outputpath & file.Name) Then
                DataImportForm.TextOutput.AppendText(String.Concat(file.Name, " was successfully converted!", vbCrLf))
            End If
        Next
    End Sub


    <CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2100:Review SQL queries for security vulnerabilities")>
    Private Function GetTable(ByVal file As IO.FileInfo, ByVal tablename As String) As DataTable
        Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & file.FullName)
        Dim command As New OleDbCommand
        Dim da As New OleDbDataAdapter
        Dim table As New DataTable(tablename)

        command.Connection = con
        command.CommandText = SQLString(tablename)
        da = New OleDbDataAdapter(command)
        con.Open()
        da.Fill(table)
        con.Close()

        GetTable = table
    End Function

    Private Function SQLString(Optional ByVal type As String = "Normal")
        If type = "Normal Table" Then
            SQLString = "SELECT * FROM Channel_Normal_Table"
        ElseIf type = "Statistics Table" Then
            SQLString = "SELECT Cycle_Index AS ""Cycle Index"", 
                              Discharge_Capacity As ""Discharge Capacity"", 
                              Charge_Capacity As ""Charge Capacity"", 
                              Discharge_Energy As ""Discharge Energy"", 
                              Charge_Energy As ""Charge Energy"" 
                            From Channel_Normal_Table 
                            INNER Join Channel_Statistic_Table 
                            On Channel_Normal_Table.Data_Point 
                            = Channel_Statistic_Table.Data_Point"
        Else
            SQLString = ""
        End If
    End Function


End Module

Module Export
    Public Function ToExcel(ByVal tables() As DataTable, ByVal filepath As String) As Boolean
        Dim xl As New Excel.Application
        Dim wbook As Excel.Workbook
        Dim wsheet As Excel.Worksheet
        Dim dt As DataTable
        Dim dc As DataColumn
        Dim dr As DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0
        Dim sheetcount As Integer = xl.SheetsInNewWorkbook


        xl.SheetsInNewWorkbook = 1
        wbook = xl.Workbooks.Add()
        wsheet = wbook.Worksheets(1)



        For Each table In tables
            wsheet.Name = table.TableName

            dt = table

            For Each dc In dt.Columns
                colIndex = colIndex + 1
                wsheet.Cells(1, colIndex) = dc.ColumnName
            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    wsheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Next
            Next

            wsheet.Columns.AutoFit()
            wsheet = wbook.Worksheets.Add()
        Next

        wbook.SaveAs(filepath)
        xl.SheetsInNewWorkbook = sheetcount
        xl.DisplayAlerts = False
        If wsheet IsNot Nothing Then If ReleaseObject(wsheet) Then wbook.Close(False)
        If ReleaseObject(wbook) Then xl.Quit()
        If ReleaseObject(xl) Then GC.Collect()

        ToExcel = True
    End Function

    Private Function ReleaseObject(ByVal o As Object) As Boolean
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
            ReleaseObject = True
        End Try
    End Function
End Module