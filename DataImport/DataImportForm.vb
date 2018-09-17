
Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports OfficeOpenXml


Public Class DataImportForm

    Protected drag As Boolean
    Protected mousex As Integer
    Protected mousey As Integer
    Protected inputdir As IO.DirectoryInfo
    Private _validexts() As String
    Protected tempfiles As New Dictionary(Of String, IO.FileInfo)
    Private _ArbinFiles As New List(Of FileInfo)()

    Public Property ArbinFiles As List(Of IO.FileInfo)
        Get
            If IsNothing(_ArbinFiles) Then
                _ArbinFiles = New List(Of FileInfo)()
            End If
            Return _ArbinFiles
        End Get
        Set
            _ArbinFiles = Value
        End Set
    End Property

    Public Sub WriteLine(ByVal text As String)
        TextOutput.AppendText(String.Format("[{0:hh:mm:ss tt}] ", Now()) & text & vbCrLf)
    End Sub

    Protected Property ValidExts As String()
        Get
            Return _validexts
        End Get
        Set(value As String())
            _validexts = value
        End Set
    End Property

    Private Sub DataImportForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        ValidExts = {".res", ".mdb"}
    End Sub

    Private Sub DataImportForm_MouseMove(sender As Object, e As EventArgs) Handles MyBase.MouseMove
        ' If drag is true, then move the form
        If drag Then
            Top = Cursor.Position.Y - mousey
            Left = Cursor.Position.X - mousex
        End If

    End Sub
    Private Sub DataImportForm_MouseDown(sender As Object, e As EventArgs) Handles MyBase.MouseDown
        ' Sets drag to true and stores mouse position
        drag = True
        mousex = Cursor.Position.X - Left
        mousey = Cursor.Position.Y - Top

    End Sub
    Private Sub DataImportForm_MouseUp(sender As Object, e As EventArgs) Handles MyBase.MouseUp
        ' Sets drag to false
        drag = False

    End Sub

    Private Sub InputBrowseButton_Click(sender As Object, e As EventArgs) Handles InputBrowseButton.Click
        If (InputFolderDialogue.ShowDialog() = DialogResult.OK) Then
            FolderInputBox.Text = InputFolderDialogue.SelectedPath
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
        If Export.IsValidFileNameOrPath(FolderInputBox.Text) Then
            inputdir = New DirectoryInfo(FolderInputBox.Text)
            FolderOutputBox.Text = FolderInputBox.Text & "\Xlsx Files\"
            PopulateFileList(inputdir)
        End If
    End Sub

    Private Sub PopulateFileList(dir As IO.DirectoryInfo)
        FileList.Items.Clear()
        For Each file In dir.EnumerateFiles()
            If Validexts.Contains(file.Extension()) Then
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
        For Each item In SelectedList.SelectedItems.OfType(Of String).ToList
            SelectedList.Items.Remove(item)
        Next
    End Sub

    Private Sub ConvertButton_Click(sender As Object, e As EventArgs) Handles ConvertButton.Click
        If IsValidFileNameOrPath(FolderOutputBox.Text) Then
            ArbinFiles.Clear()
            For Each filename In SelectedList.Items
                If {".res", ".mdb"}.Contains(tempfiles(filename).Extension) Then
                    ArbinFiles.Add(tempfiles(filename))
                End If
            Next

            WriteLine("Converting files:")
            For Each file In ArbinFiles
                TextOutput.AppendText(String.Concat(file.Name, vbCrLf))
            Next

            Arbin.Convert(ArbinFiles, FolderOutputBox.Text)
        Else
            MsgBox("Output file path is not valid.")
            WriteLine("Invalid output directory.")

        End If

        WriteLine("Conversion completed!")
    End Sub

End Class

Module Arbin
    Public Sub Convert(ByVal files As List(Of IO.FileInfo), outputpath As String)
        Dim ds As New DataSet
        Dim outputdir As New IO.DirectoryInfo(outputpath)
        Dim timeperfile As New Stopwatch

        For Each file In files
            timeperfile = Stopwatch.StartNew()
            DataImportForm.WriteLine("Pulling tables for " & file.Name & ".")

            ds = GetDataSet(file)

            DataImportForm.WriteLine("Exporting to Excel.")
            If Export.ToExcel(ds, outputpath & "\" & IO.Path.GetFileNameWithoutExtension(file.Name) & ".xlsx") Then
                DataImportForm.WriteLine(String.Format("{0} was successfully converted in {1} s!", file.Name, timeperfile.ElapsedMilliseconds / 1000))
            End If
        Next
    End Sub

    Private Function GetDataSet(ByVal file As IO.FileInfo) As DataSet
        Dim ds As New DataSet
        Dim Tables As New Dictionary(Of String, Dictionary(Of String, Type))
        Dim normalcols As New Dictionary(Of String, Type)
        Dim statcols As New Dictionary(Of String, Type)



        With normalcols
            .Add("Test_ID", GetType(Int32))
            .Add("Data_Point", GetType(Int64))
            .Add("Test_Time", GetType(Double))
            .Add("Step_Time", GetType(Double))
            .Add("DateTime", GetType(Double))
            .Add("Step_Index", GetType(Int32))
            .Add("Cycle_Index", GetType(Int32))
            .Add("Is_FC_Data", GetType(Boolean))
            .Add("Current", GetType(Double))
            .Add("Voltage", GetType(Double))
            .Add("Charge_Capacity", GetType(Double))
            .Add("Discharge_Capacity", GetType(Double))
            .Add("Charge_Energy", GetType(Double))
            .Add("Discharge_Energy", GetType(Double))
            .Add("dV/dt", GetType(Double))
            .Add("Internal_Resistance", GetType(Double))
            .Add("AC_Impedance", GetType(Double))
            .Add("ACI_Phase_Angle", GetType(Double))
        End With

        With statcols
            .Add("Cycle_Index", GetType(Int32))
            .Add("Discharge_Capacity", GetType(Double))
            .Add("Charge_Capacity", GetType(Double))
            .Add("Discharge_Energy", GetType(Double))
            .Add("Charge_Energy", GetType(Double))
        End With

        Tables.Add("Normal Table", normalcols)
        Tables.Add("Statistics Table", statcols)

        ' Creates two tables in a dataset 
        For Each tablename In Tables.Keys()
            ds.Tables.Add(tablename)
            For Each colname In Tables(tablename).Keys()
                ds.Tables(tablename).Columns.Add(colname, Tables.Item(tablename).Item(colname))
            Next
        Next

        ' Fill DataSet from file
        FillDataSet(file, ds)

        Return ds


    End Function

    <CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2100:Review SQL queries for security vulnerabilities")>
    Private Function FillDataSet(ByVal file As FileInfo, ByRef ds As DataSet)
        Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & file.FullName)
        Dim command As New OleDbCommand
        Dim da As New OleDbDataAdapter


        command.Connection = con
        con.Open()
        For Each dt In ds.Tables
            command.CommandText = SQLString(dt.TableName)
            da = New OleDbDataAdapter(command)
            da.Fill(dt)
        Next
        con.Close()
    End Function

    Private Function SQLString(Optional ByVal type As String = "Normal")
        If type = "Normal Table" Then
            SQLString = "SELECT * FROM Channel_Normal_Table ORDER BY Data_Point ASC"
        ElseIf type = "Statistics Table" Then
            SQLString = "SELECT Cycle_Index, 
                              Discharge_Capacity, 
                              Charge_Capacity, 
                              Discharge_Energy, 
                              Charge_Energy
                            From Channel_Normal_Table 
                            INNER Join Channel_Statistic_Table 
                            On Channel_Normal_Table.Data_Point 
                            = Channel_Statistic_Table.Data_Point
                            Order BY Cycle_Index ASC"
        Else
            SQLString = ""
        End If
    End Function


End Module

Module Export
    Public Function ToExcel(ByVal ds As DataSet, ByVal filepath As String) As Boolean
        Dim dt As DataTable
        Dim ws As ExcelWorksheet
        Dim file As New FileInfo(filepath)
        Dim colnumber As Int32


        Using pck = New ExcelPackage
            For Each dt In ds.Tables()
                ws = pck.Workbook.Worksheets.Add(dt.TableName)
                ws.Cells("A1").LoadFromDataTable(dt, True)

                colnumber = 1
                For Each col In dt.Columns
                    ws.Cells(1, colnumber).Value = ws.Cells(1, colnumber).Value.Replace("_", " ")

                    If col.ColumnName = "DateTime" Then
                        ws.Column(colnumber).Style.Numberformat.Format = "MM/dd/yyyy hh:mm:ss AM/PM"
                    End If
                    colnumber += 1
                Next
            Next
            Try
                pck.SaveAs(file)
            Catch
                DataImportForm.WriteLine(String.Format("Error saving file: {0}.", file.Name))
                Return False
            End Try
        End Using

        Return True

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
    Function IsValidFileNameOrPath(ByVal name As String,
                                   Optional ByVal createnew As Boolean = True,
                                   Optional ByVal readonlyvalid As Boolean = False,
                                   Optional ByVal isfile As Boolean = False) As Boolean
        Dim di As DirectoryInfo
        Dim fi As FileInfo

        If name.Contains("//") Then
            Return False
        End If

        ' Determines if the name is Nothing.
        If name Is Nothing Then
            Return False
        End If



        ' Determines if the directory exists and tries to create
        If Not isfile Then
            ' Determines if there are bad characters in the name.
            For Each badChar As Char In Path.GetInvalidPathChars
                If InStr(name, badChar) > 0 Then
                    Return False
                End If
            Next

            di = New DirectoryInfo(name)
            If Not di.Exists() Then
                If createnew Then
                    Try
                        di.Create()
                        di.Refresh()
                        di.Attributes = di.Attributes Or FileAttributes.Normal
                    Catch
                        Return False
                    End Try
                Else
                    Return False
                End If
            End If

            ' Determines if file is read-only
            If Not readonlyvalid Then
                If di.Attributes.HasFlag(FileAttributes.ReadOnly) Then
                    Return False
                End If
            End If
        Else
            fi = New FileInfo(name)
            If Not readonlyvalid Then
                If fi.Exists() Then
                    If fi.Attributes.HasFlag(FileAttributes.ReadOnly) Then
                        Return False
                    End If
                End If
            End If
        End If



        ' The name passes basic validation.
        Return True
    End Function
End Module