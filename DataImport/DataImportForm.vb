
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.IO
Imports System.Reflection
Imports Microsoft.VisualBasic.FileIO
Imports OfficeOpenXml
Imports Squirrel
Imports Excel = Microsoft.Office.Interop.Excel


Public Class DataImportForm
    Protected drag As Boolean
    Protected mousex As Integer
    Protected mousey As Integer
    Protected inputdir As IO.DirectoryInfo
    Private _validexts() As String
    Protected tempfiles As New Dictionary(Of String, IO.FileInfo)
    Private _ArbinFiles As New List(Of FileInfo)()
    Private _EISFiles As New List(Of FileInfo)()





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

    Public Property EISFiles As List(Of IO.FileInfo)
        Get
            If IsNothing(_EISFiles) Then
                _EISFiles = New List(Of FileInfo)()
            End If
            Return _EISFiles
        End Get
        Set
            _EISFiles = Value
        End Set
    End Property


    Protected Property ValidExts As String()
        Get
            Return _validexts
        End Get
        Set(value As String())
            _validexts = value
        End Set
    End Property

    Private Async Sub DataImportForm_LoadAsync(sender As Object, e As EventArgs) Handles Me.Load
        ValidExts = {".res", ".mdb", ".par"}

        textVersionNumber.Text = "Version " & Assembly.GetExecutingAssembly().GetName().Version.ToString()

        Using mgr = New UpdateManager("G:\Global\Transfer\Internal\akleiderer\Releases")
            Await mgr.UpdateApp()
        End Using

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

    Private Sub PopulateFileList(dir As DirectoryInfo)
        '#TODO Bug occurs when same file name is in two different folders.

        FileList.Items.Clear()
        For Each file In dir.EnumerateFiles()
            If ValidExts.Contains(file.Extension()) Then
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
        BackgroundWorker1.RunWorkerAsync()
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim i As Int32 = 0
        If IsValidFileNameOrPath(FolderOutputBox.Text) Then
            ArbinFiles.Clear()
            EISFiles.Clear()
            For Each filename In SelectedList.Items
                If {".res", ".mdb"}.Contains(tempfiles(filename).Extension) Then
                    ArbinFiles.Add(tempfiles(filename))
                End If
                If {".par"}.Contains(tempfiles(filename).Extension) Then
                    EISFiles.Add(tempfiles(filename))
                End If
            Next
            BackgroundWorker1.ReportProgress(i, New String("Converting files:"))
            For Each file In ArbinFiles
                BackgroundWorker1.ReportProgress(i, New String("- " & file.Name))
            Next
            For Each file In ArbinFiles
                Dim ds As New DataSet
                Dim outputdir As New IO.DirectoryInfo(FolderOutputBox.Text)
                Dim timeperfile As New Stopwatch
                Dim xlsApp As Excel.Application = Nothing
                Dim xlsWorkBooks As Excel.Workbooks = Nothing
                Dim xlsWB As Excel.Workbook = Nothing


                timeperfile = Stopwatch.StartNew()
                BackgroundWorker1.ReportProgress(0, New String("Pulling tables for " & file.Name & "."))

                ds = Arbin.GetDataSet(file)

                BackgroundWorker1.ReportProgress(0, New String("Exporting to Excel."))
                If Export.ToExcel(ds, FolderOutputBox.Text & "\" & IO.Path.GetFileNameWithoutExtension(file.Name) & ".xlsx") Then
                    Try
                        xlsApp = New Excel.Application
                        xlsApp.Visible = False
                        xlsWorkBooks = xlsApp.Workbooks
                        xlsWB = xlsWorkBooks.Open(FolderOutputBox.Text & "\" & IO.Path.GetFileNameWithoutExtension(file.Name) & ".xlsx")

                        ArbinFormatMain(xlsWB)
                        xlsWB.Save()

                    Catch ex As Exception
                        MessageBox.Show(String.Format("Error: {0}", ex.Message))

                    Finally
                        xlsWB.Close()
                        xlsWB = Nothing
                        xlsApp.Quit()
                        xlsApp = Nothing
                    End Try
                    BackgroundWorker1.ReportProgress(0, New String(String.Format("{0} was successfully converted in {1} s!", file.Name, timeperfile.ElapsedMilliseconds / 1000)))
                End If

            Next

            For Each file In EISFiles
                BackgroundWorker1.ReportProgress(i, New String("- " & file.Name))
            Next
            For Each file In EISFiles
                Dim ds As New DataSet
                Dim outputdir As New IO.DirectoryInfo(FolderOutputBox.Text)
                Dim timeperfile As New Stopwatch
                Dim xlsApp As Excel.Application = Nothing
                Dim xlsWorkBooks As Excel.Workbooks = Nothing
                Dim xlsWB As Excel.Workbook = Nothing


                timeperfile = Stopwatch.StartNew()
                BackgroundWorker1.ReportProgress(0, New String("Pulling tables for " & file.Name & "."))

                ds = EIS.EISGetDataSet(file)

                BackgroundWorker1.ReportProgress(0, New String("Exporting to Excel."))
                If Export.ToExcel(ds, FolderOutputBox.Text & "\" & IO.Path.GetFileNameWithoutExtension(file.Name) & "_EIS.xlsx") Then
                    Try
                        xlsApp = New Excel.Application
                        xlsApp.Visible = False
                        xlsWorkBooks = xlsApp.Workbooks
                        xlsWB = xlsWorkBooks.Open(FolderOutputBox.Text & "\" & IO.Path.GetFileNameWithoutExtension(file.Name) & "_EIS.xlsx")

                        EISFormatMain(xlsWB)
                        xlsWB.Save()

                    Catch ex As Exception
                        MessageBox.Show(String.Format("Error: {0}", ex.Message))

                    Finally
                        xlsWB.Close()
                        xlsWB = Nothing
                        xlsApp.Quit()
                        xlsApp = Nothing
                    End Try
                    BackgroundWorker1.ReportProgress(0, New String(String.Format("{0} was successfully converted in {1} s!", file.Name, timeperfile.ElapsedMilliseconds / 1000)))
                End If

            Next

        Else
            MsgBox("Output file path is not valid.")
            BackgroundWorker1.ReportProgress(i, New String("Invalid output directory."))
        End If
        BackgroundWorker1.ReportProgress(i, New String("Conversion completed!"))
    End Sub

    Public Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        TextOutput.AppendText(String.Format("[{0:hh:mm:ss tt}] ", Now()) & e.UserState & vbCrLf)
    End Sub


End Class

Module Arbin

    Public Function GetDataSet(ByVal file As IO.FileInfo) As DataSet
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

        Return True
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

Module EIS
    Public Function EISGetDataSet(ByVal file As IO.FileInfo) As DataSet
        Dim ds As New DataSet
        Dim Tables As New Dictionary(Of String, Dictionary(Of String, Type))
        Dim cols As New Dictionary(Of String, Type)
        Dim statcols As New Dictionary(Of String, Type)

        ds.Tables.Add("EIS_Data")

        With ds.Tables("EIS_Data").Columns
            .Add("Segment", GetType(Int32))
            .Add("Point", GetType(Int64))
            .Add("Voltage", GetType(Double))
            .Add("Current", GetType(Double))
            .Add("Elapsed_Time", GetType(Double))
            .Add("ADC_Sync_Input", GetType(Int32))
            .Add("Current_Range", GetType(Int32))
            .Add("Status", GetType(Int32))
            .Add("Voltage_Applied", GetType(Double))
            .Add("Frequency", GetType(Double))
            .Add("Voltage_Real", GetType(Double))
            .Add("Voltage_Imag", GetType(Double))
            .Add("Impedance_Real", GetType(Double))
            .Add("Impedance_Imag", GetType(Double))
            .Add("Z_Real", GetType(Double))
            .Add("Z_Imag", GetType(Double))
            .Add("Voltage2_Status", GetType(Int32))
            .Add("Voltage2", GetType(Double))
            .Add("Voltage2_Real", GetType(Double))
            .Add("Voltage2_Imag", GetType(Double))
            .Add("Z2_Real", GetType(Double))
            .Add("Z2_Imag", GetType(Double))
            .Add("ActionID", GetType(Int32))
            .Add("AC_Amplitude", GetType(Double))
        End With

        ' Fill DataSet from file
        EISFillDataSet(file, ds)
        Return ds
    End Function

    Public Function EISFillDataSet(ByVal file As FileInfo, ByRef ds As DataSet)

        Using MyReader As New Microsoft.VisualBasic.
                    FileIO.TextFieldParser(
                    file.FullName)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            Dim currentRow As String()
            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    If currentRow.Length = 24 Then
                        ds.Tables("EIS_Data").Rows.Add(currentRow)
                    End If
                Catch ex As Microsoft.VisualBasic.
                        FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message &
                "is not valid and will be skipped.")
                End Try
            End While
        End Using

        Return True
    End Function

End Module

Module ExcelFormat
    Public Sub ArbinFormatMain(ByRef wb As Excel.Workbook)

        Dim ws As Excel.Worksheet

        ws = wb.Worksheets(1)

        ' Creates the Cell Info Table
        If Not CellInfo(ws) Then
            Exit Sub
        End If

        ' Make data table
        ' Insert specific capacity/energy columns
        If Not ArbinDataTable(ws) Then
            Exit Sub
        End If

        ' Make graph
        ' Format graph
        If Not ArbinDataChart(ws) Then
            Exit Sub
        End If


    End Sub

    Public Sub EISFormatMain(ByRef wb As Excel.Workbook)
        Dim ws As Excel.Worksheet

        ws = wb.Worksheets(1)

        ' Creates the Cell Info Table
        If Not CellInfo(ws) Then
            Exit Sub
        End If

        ' Make data table
        If Not EISDataTable(ws) Then
            Exit Sub
        End If

        ' Make graph
        ' Format graph
        If Not EISDataChart(ws) Then
            Exit Sub
        End If

    End Sub

    Private Function CellInfo(ByRef ws As Excel.Worksheet)
        Dim objTable As Excel.ListObject

        Try
            ws.Rows("1:3").Insert(Shift:=Excel.XlDirection.xlDown, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
            ws.Range("A1").Value = "Cell Name"
            ws.Range("B1").Value = "Cathode Weight (mg)"
            ws.Range("C1").Value = "Active Weight (mg)"

            objTable = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, ws.Range("A1:C2"), , Excel.XlYesNoGuess.xlYes,)
            objTable.TableStyle = "TableStyleLight1"
            objTable.DisplayName = "Cell_Info"
            objTable.Name = "Cell_Info"

            Return True

        Catch
            MsgBox("There was an error creating the Cell Info table.")
            Return False

        End Try

    End Function

    Private Function ArbinDataTable(ByRef ws As Excel.Worksheet) As Boolean
        Dim objTable As Excel.ListObject
        Dim rng As Excel.Range

        Try
            rng = ws.Range("A4", ws.Range("A4").End(Excel.XlDirection.xlDown).End(Excel.XlDirection.xlToRight))

            objTable = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rng, , Excel.XlYesNoGuess.xlYes, )

            With objTable

                .DisplayName = "Data_Table"
                .Name = "Data_Table"
                .TableStyle = "TableStyleLight1"
                .ListColumns.Add(13)
                .HeaderRowRange(13) = "Specific Discharge Capacity"
                .ListColumns.Add(16)
                .HeaderRowRange(16) = "Specific Discharge Energy"
                .ListColumns(13).DataBodyRange.FormulaR1C1 = "=[Discharge Capacity]*1000000/Cell_Info[Active Weight (mg)]"
                .ListColumns(16).DataBodyRange.FormulaR1C1 = "=[Discharge Energy]*1000000/Cell_Info[Active Weight (mg)]"

            End With

            Return True

        Catch
            MsgBox("There was an error creating the Data table.")
            Return False

        End Try

    End Function

    Private Function ArbinDataChart(ByRef ws As Excel.Worksheet) As Boolean
        Dim objChart As Object
        Dim objTable As Excel.ListObject
        Dim objSeries As Excel.Series

        Try
            objTable = ws.ListObjects("Data_Table")

            objTable.Range.AutoFilter(Field:=6, Criteria1:="2")
            objTable.Range.AutoFilter(Field:=7, Criteria1:="1")

            objChart = ws.ChartObjects.Add(
                Left:=ws.Range("A5").Left,
                Width:=ws.Application.InchesToPoints(6),
                Top:=ws.Range("A5").Top,
                Height:=ws.Application.InchesToPoints(3.6))

            With objChart.Chart

                .ChartType = Excel.XlChartType.xlXYScatter
                objSeries = .SeriesCollection.NewSeries()

                ' Add Data
                With objSeries
                    .Name = "='Normal Table'!$A$2"
                    .XValues = objTable.ListColumns("Specific Discharge Capacity").DataBodyRange
                    .Values = objTable.ListColumns("Voltage").DataBodyRange
                End With

                .ChartTitle.Delete
                .HasTitle = False
                .HasLegend = False

                'Format Y Axis
                With .Axes(Excel.XlAxisType.xlValue)
                    .MajorGridlines.Delete
                    .MinorGridlines.Delete
                    .HasTitle = True
                    .AxisTitle.Text = "Voltage (V)"

                    With .TickLabels.Font
                        .Name = "Arial"
                        .Size = 10
                        .Color = RGB(0, 0, 0)
                    End With

                    With .AxisTitle.Font
                        .Name = "Arial"
                        .Size = 11
                        .Bold = True
                    End With
                End With

                'Format X Axis
                With .Axes(Excel.XlAxisType.xlCategory)
                    .MajorGridlines.Delete
                    .MinorGridlines.Delete
                    .HasTitle = True
                    .AxisTitle.Text = "Specific Discharge Capacity (mAh/g)"

                    With .TickLabels.Font
                        .Name = "Arial"
                        .Size = 10
                        .Color = RGB(0, 0, 0)
                    End With

                    With .AxisTitle.Font
                        .Name = "Arial"
                        .Size = 11
                        .Bold = True
                    End With

                End With

            End With

            Return True

        Catch
            MsgBox("There was an error creating the Data chart.")
            Return False

        End Try

    End Function

    Private Function EISDataTable(ByRef ws As Excel.Worksheet) As Boolean
        Dim objTable As Excel.ListObject
        Dim rng As Excel.Range

        Try
            rng = ws.Range("A4", ws.Range("A4").End(Excel.XlDirection.xlDown).End(Excel.XlDirection.xlToRight))

            objTable = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rng, , Excel.XlYesNoGuess.xlYes, )

            With objTable
                .DisplayName = "Data_Table"
                .Name = "Data_Table"
                .TableStyle = "TableStyleLight1"
            End With

            Return True

        Catch
            MsgBox("There was an error creating the Data table.")
            Return False

        End Try

    End Function

    Private Function EISDataChart(ByRef ws As Excel.Worksheet) As Boolean
        Dim objChart As Object
        Dim objTable As Excel.ListObject
        Dim objSeries As Excel.Series

        Try
            objTable = ws.ListObjects("Data_Table")

            objChart = ws.ChartObjects.Add(
                Left:=ws.Range("A5").Left,
                Width:=ws.Application.InchesToPoints(6),
                Top:=ws.Range("A5").Top,
                Height:=ws.Application.InchesToPoints(3.6))

            With objChart.Chart

                .ChartType = Excel.XlChartType.xlXYScatter
                objSeries = .SeriesCollection.NewSeries()

                ' Add Data
                With objSeries
                    .Name = "='EIS_Data'!$A$2"
                    .XValues = objTable.ListColumns("Z Real").DataBodyRange
                    .Values = objTable.ListColumns("Z Imag").DataBodyRange
                    .MarkerStyle = 8
                    .MarkerSize = 5
                    .Format.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                End With

                .ChartTitle.Delete
                .HasTitle = False
                .HasLegend = False

                'Format Y Axis
                With .Axes(Excel.XlAxisType.xlValue)
                    .MajorGridlines.Delete
                    .MinorGridlines.Delete
                    .ReversePlotOrder = True
                    .HasTitle = True
                    .AxisTitle.Text = "Z Imaginary (" & ChrW(937).ToString & ")"
                    .TickLabels.NumberFormat = "#,##0;#,##0"

                    With .TickLabels.Font
                        .Name = "Arial"
                        .Size = 10
                        .Color = RGB(0, 0, 0)
                    End With

                    With .AxisTitle.Font
                        .Name = "Arial"
                        .Size = 11
                        .Bold = True
                    End With
                End With

                .PlotArea.Left = 40
                .PlotArea.Top = 5

                'Format X Axis
                With .Axes(Excel.XlAxisType.xlCategory)
                    .MajorGridlines.Delete
                    .MinorGridlines.Delete
                    .HasTitle = True
                    .AxisTitle.Text = "Z Real (" & ChrW(937).ToString & ")"
                    .AxisTitle.Left = 201.937
                    .AxisTitle.Top = 240
                    .TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionHigh

                    With .TickLabels.Font
                        .Name = "Arial"
                        .Size = 10
                        .Color = RGB(0, 0, 0)
                    End With

                    With .AxisTitle.Font
                        .Name = "Arial"
                        .Size = 11
                        .Bold = True
                    End With

                End With



            End With

            Return True

        Catch
            MsgBox("There was an error creating the Data chart.")
            Return False

        End Try

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
                DataImportForm.BackgroundWorker1.ReportProgress(0, New String(String.Format("Error saving file: {0}.", file.Name)))
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