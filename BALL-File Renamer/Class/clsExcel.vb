Imports Excel = Microsoft.Office.Interop.Excel
Imports System
Imports System.IO

Public Class clsExcel
    Public Shared Sub LoadExcelDataToDataGrid(filepath As String, dgv As DataGridView)


        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim DtSet As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\Users\Administrator\Desktop\VB_PROJECTS\BALL\Ball\Ball\Resources\PLC1.xls';Extended Properties=Excel 8.0;")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\Users\Administrator\Desktop\VB_PROJECTS\BALL\Ball\Ball\Resources\PLC1.xls';Extended Properties=Excel 8.0;")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & filepath & "';Extended Properties=Excel 8.0;")
        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & filepath & "';Extended Properties='Excel 8.0;IMEX=1;';")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & System.AppDomain.CurrentDomain.BaseDirectory.ToString() & "PLC1.xls';Extended Properties=Excel 8.0;")

        'MsgBox("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & System.AppDomain.CurrentDomain.BaseDirectory.ToString() & "PLC1.xls';Extended Properties=Excel 8.0;")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='\PLC1.xls';Extended Properties=Excel 8.0;")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\Users\Administrator\Desktop\VB_PROJECTS\BALL\Ball\Ball\bin\Debug\PLC1.xls';Extended Properties=Excel 8.0;")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & Application.StartupPath & "\PLC1.xls';Extended Properties=Excel 8.0;")
        'MsgBox(Application.StartupPath & "\PLC1.xls'")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='|DataDirectory|PLC1.xls';Extended Properties=Excel 8.0;")



        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
        'MyCommand.TableMappings.Add("Table", "Net-informations.com")
        DtSet = New System.Data.DataSet

        MyCommand.Fill(DtSet)

        'DtSet.Tables(0).Columns.Add(EXCEL_HEADER_currentValue, Type.GetType("System.String"))

        dgv.DataSource = DtSet.Tables(0)
        MyConnection.Close()
    End Sub


    Public Shared Function LoadExcelDataToDataSet(filepath As String, strQuery As String)


        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim DtSet As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\Users\Administrator\Desktop\VB_PROJECTS\BALL\Ball\Ball\Resources\PLC1.xls';Extended Properties=Excel 8.0;")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\Users\Administrator\Desktop\VB_PROJECTS\BALL\Ball\Ball\Resources\PLC1.xls';Extended Properties=Excel 8.0;")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & filepath & "';Extended Properties=Excel 8.0;")
        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & filepath & "';Extended Properties='Excel 8.0;IMEX=1;';")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & System.AppDomain.CurrentDomain.BaseDirectory.ToString() & "PLC1.xls';Extended Properties=Excel 8.0;")

        'MsgBox("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & System.AppDomain.CurrentDomain.BaseDirectory.ToString() & "PLC1.xls';Extended Properties=Excel 8.0;")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='\PLC1.xls';Extended Properties=Excel 8.0;")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\Users\Administrator\Desktop\VB_PROJECTS\BALL\Ball\Ball\bin\Debug\PLC1.xls';Extended Properties=Excel 8.0;")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & Application.StartupPath & "\PLC1.xls';Extended Properties=Excel 8.0;")
        'MsgBox(Application.StartupPath & "\PLC1.xls'")
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='|DataDirectory|PLC1.xls';Extended Properties=Excel 8.0;")



        'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
        MyCommand = New System.Data.OleDb.OleDbDataAdapter(strQuery, MyConnection)
        'MyCommand.TableMappings.Add("Table", "Net-informations.com")
        DtSet = New System.Data.DataSet

        MyCommand.Fill(DtSet)



        'dgv.DataSource = DtSet.Tables(0)
        MyConnection.Close()

        Return DtSet
    End Function


    Public Shared Function listToExcel(ByVal strFilename As String, ByVal strSheetname As String, ByVal strListIn As List(Of String)) As Integer

        Dim strListList_Temp As List(Of List(Of String)) = New List(Of List(Of String))()
        strListList_Temp.Add(strListIn)

        Return listToExcel(strFilename, strSheetname, strListList_Temp)


    End Function

    Public Shared Function listToExcel(ByVal strFilename As String, ByVal strSheetname As String, ByVal strListListIn As List(Of List(Of String))) As Integer
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet = Nothing


        Try




            If xlApp Is Nothing Then
                MessageBox.Show("Excel is not properly installed!!")
                Return -1
            End If




            If IO.File.Exists(strFilename) Then
                xlWorkBook = xlApp.Workbooks.Open(strFilename)
            Else
                xlWorkBook = xlApp.Workbooks.Add()
            End If


            If DoesSheetExists(xlWorkBook, strSheetname) Then
                xlWorkSheet = xlWorkBook.Sheets(strSheetname)
                xlWorkSheet.Cells.Delete()
            Else
                xlWorkSheet = xlWorkBook.Sheets.Add()
                'xlWorkSheet.Name = "ren"
                xlWorkSheet.Name = strSheetname
            End If








            'xlWorkSheet.Cells(1, 1) = "Filenames 1"
            'xlWorkSheet.Cells(1, 2) = "Filenames 2"

            Dim intRow As Integer
            Dim intCol As Integer = 1
            For Each strList In strListListIn
                intRow = 1
                For Each strTemp In strList
                    xlWorkSheet.Cells(intRow, intCol) = strTemp
                    'xlWorkSheet.Cells(intRow, 2) = Path.Combine(Path.GetDirectoryName(strTemp), (intRow - 2).ToString & ".doc")
                    intRow = intRow + 1
                Next
                intCol = intCol + 1
            Next



            'xlWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden


            'xlWorkSheet = xlWorkBook.Sheets.Add()
            'xlWorkSheet.Name = "(H)"

            'xlWorkBook.SaveAs("d:\csharp-Excel.xls")
            'xlWorkBook.SaveAs("D:\" & strFilename & ".xls", Excel.XlFileFormat.xlExcel8)

            'While File.Exists(strFilename) AndAlso IsFileOpen(strFilename)
            '    MsgBox(strFilename & " is currently open, please close it to continue.")
            'End While
            xlApp.DisplayAlerts = False

            xlWorkBook.SaveAs(strFilename, Excel.XlFileFormat.xlExcel8)





            'xlWorkBook.SaveAs("d:\csharp-Excel.xls")
            'xlWorkBook.SaveAs("D:\" & strExcelFilename & ".xls", Excel.XlFileFormat.xlExcel8)
            'xlWorkBook.Save()


            xlWorkBook.Close()
            xlApp.Quit()

            xlApp.DisplayAlerts = True

            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)
            Return 100
        Catch ex As Exception


            xlWorkBook.Close()
            xlApp.Quit()

            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)


        End Try
        Return -1
    End Function

    Public Shared Function listToExcelOld(ByVal strFilename As String, ByVal strListListIn As List(Of List(Of String))) As Integer
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
            Return -1
        End If

        Dim xlWorkBook As Excel.Workbook
        xlWorkBook = xlApp.Workbooks.Add()
        'xlWorkBook = xlApp.Workbooks.Open(strFilename)



        Dim xlWorkSheet As Excel.Worksheet
        xlWorkSheet = xlWorkBook.Sheets.Add()
        'xlWorkSheet.Name = "ren"
        xlWorkSheet.Name = varFromJson.Item("strings").Item("worksheet.name")


        'xlWorkSheet.Cells(1, 1) = "Filenames 1"
        'xlWorkSheet.Cells(1, 2) = "Filenames 2"

        Dim intRow As Integer
        Dim intCol As Integer = 1
        For Each strList In strListListIn
            intRow = 1
            For Each strTemp In strList
                xlWorkSheet.Cells(intRow, intCol) = strTemp
                'xlWorkSheet.Cells(intRow, 2) = Path.Combine(Path.GetDirectoryName(strTemp), (intRow - 2).ToString & ".doc")
                intRow = intRow + 1
            Next
            intCol = intCol + 1
        Next



        xlWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden


        xlWorkSheet = xlWorkBook.Sheets.Add()
        xlWorkSheet.Name = "(H)"

        'xlWorkBook.SaveAs("d:\csharp-Excel.xls")
        'xlWorkBook.SaveAs("D:\" & strFilename & ".xls", Excel.XlFileFormat.xlExcel8)

        While File.Exists(strFilename) AndAlso IsFileOpen(strFilename)
            MsgBox(strFilename & " is currently open, please close it to continue.")
        End While

        xlWorkBook.SaveAs(strFilename, Excel.XlFileFormat.xlExcel8)





        'xlWorkBook.SaveAs("d:\csharp-Excel.xls")
        'xlWorkBook.SaveAs("D:\" & strExcelFilename & ".xls", Excel.XlFileFormat.xlExcel8)
        'xlWorkBook.Save()

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        Return 100
    End Function


    Public Shared Function DoesSheetExists(ByVal xlWorkBook As Excel.Workbook, ByVal shtName As String) As Boolean
        Dim xs As Excel.Worksheet

        DoesSheetExists = False

        '~~> Loop through the all the sheets in the workbook to find if name matches
        For Each xs In xlWorkBook.Sheets
            If xs.Name = shtName Then
                DoesSheetExists = True
            End If
        Next
    End Function

    Public Shared Function IsFileOpen(ByRef sName As String) As Boolean
        Dim blnRetVal As Boolean = False
        Dim fs As FileStream = Nothing

        Try
            fs = File.Open(sName, FileMode.Open, FileAccess.Read, FileShare.None)
        Catch ex As Exception
            blnRetVal = True
        Finally
            If Not IsNothing(fs) Then
                fs.Close()
            End If

        End Try
        Return blnRetVal

    End Function

    Public Shared Function CreateExcelFile(ByVal strExcelFilename As String) As Excel.Workbook
        If IO.File.Exists(strExcelFilename) Then
            Return Nothing
        End If
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
            Return Nothing
        End If

        Dim xlWorkBook As Excel.Workbook
        xlWorkBook = xlApp.Workbooks.Add()



        Dim xlWorkSheet As Excel.Worksheet
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        xlWorkSheet.Cells(1, 1) = "Sheet 1 content"

        'xlWorkBook.SaveAs("d:\csharp-Excel.xls")
        xlApp.DisplayAlerts = False
        xlWorkBook.SaveAs("D:\" & strExcelFilename & ".xls", Excel.XlFileFormat.xlExcel8)

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        Return xlWorkBook
    End Function

    Public Shared Function CreateExcelFileOLD(ByVal strExcelFilename As String) As Integer

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
            Return -1
        End If

        Dim xlWorkBook As Excel.Workbook
        xlWorkBook = xlApp.Workbooks.Add()



        Dim xlWorkSheet As Excel.Worksheet
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        xlWorkSheet.Cells(1, 1) = "Sheet 1 content"

        'xlWorkBook.SaveAs("d:\csharp-Excel.xls")
        xlWorkBook.SaveAs("D:\" & strExcelFilename & ".xls", Excel.XlFileFormat.xlExcel8)

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        Return 100
    End Function



    Private Shared Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class
