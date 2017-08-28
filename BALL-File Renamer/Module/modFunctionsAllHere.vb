Imports System
Imports System.IO

Module modFunctionsAllHere



    Sub RenameFileBasedOnExcel(ByVal strSheetname As String)
        Try
            RefreshJson()

            Dim DtSet As System.Data.DataSet
            DtSet = clsExcel.LoadExcelDataToDataSet(StrMakeFullPath(varFromJson.Item("filenames").Item("excel4rename")), "select * from [" & strSheetname & "$]")

            If DtSet.Tables(0).Rows.Count = 0 Then

                MsgBox("There are no pair of file!")
                Exit Sub

            End If

            Dim intAllRow As Integer = DtSet.Tables(0).Rows.Count
            Dim intFileExistsInCol0 As Integer = 0
            Dim intFileExistsInCol1 As Integer = 0


            'Check which side is current
            For Each row As DataRow In DtSet.Tables(0).Rows
                If File.Exists(row.Item(0)) Or Directory.Exists(row.Item(0)) Then
                    intFileExistsInCol0 = intFileExistsInCol0 + 1
                End If

                If File.Exists(row.Item(1)) Or Directory.Exists(row.Item(1)) Then
                    intFileExistsInCol1 = intFileExistsInCol1 + 1
                End If
            Next

            Dim oldCol, newCol As Integer

            If intFileExistsInCol0 = intAllRow Then
                oldCol = 0
                newCol = 1
            ElseIf intFileExistsInCol1 = intAllRow Then
                oldCol = 1
                newCol = 0
            Else
                MsgBox("Neither column has all exists filenames")

                Exit Sub

            End If


            Dim strError As String = ""

            For Each row As DataRow In DtSet.Tables(0).Rows
                For Each row2 As DataRow In DtSet.Tables(0).Rows
                    If row.Item(0) = row2.Item(1) Then
                        MsgBox(String.Format("At least one filename appear in both column, eg.{1}{0}", row.Item(0), Environment.NewLine))
                        Exit Sub
                    End If
                Next

            Next



            For Each row As DataRow In DtSet.Tables(0).Rows
                If Directory.Exists(row.Item(oldCol)) Then
                    My.Computer.FileSystem.MoveDirectory(row.Item(oldCol), row.Item(newCol))
                ElseIf File.Exists(row.Item(oldCol)) Then
                    My.Computer.FileSystem.MoveFile(row.Item(oldCol), row.Item(newCol))


                End If
                'My.Computer.FileSystem.RenameFile(row.Item(oldCol), Path.GetFileName(row.Item(newCol)))


            Next

            MsgBox(String.Format("All files are renamed (Column {0} -> {1})", oldCol, newCol))

        Catch ex As Exception
            MsgBox(GetExceptionInfo(ex))
        End Try
    End Sub


    Enum PrepareFileForRenameMode
        FilesAndFoldersInFirstLevel
        FilesInAllSubDir
    End Enum
    Sub PrepareFileForRename(ByVal mode As String)
        Try
            RefreshJson()

            'Dim pathIn As String = InputBox("input main path here")
            'Dim pathIn As String = FileOpenDialog()

            frmPleaseWait1.Show()
            frmPleaseWait1.TextBox1.Text = ""
            frmPleaseWait1.ProgressBar1.Value = 0
            ' Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor

            Application.DoEvents()

            ' Execute your GoToSheets method here


            Dim strOutputFile As String = varFromJson.Item("filenames").Item("excel4rename")



            Dim pathIn As String = ""

            Dim openFileDialog1 As New OpenFileDialog()


            openFileDialog1.Title = "Choose Folder by select one of its child."

            'openFileDialog1.InitialDirectory = "D:\cant rename"
            openFileDialog1.InitialDirectory = Environment.CurrentDirectory
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True
            'openFileDialog1.a()

            If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                'pathIn = openFileDialog1.FileName

                pathIn = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.LastIndexOf("\"))

                'With openFileDialog1

                '    pathIn = .FileName.Substring(0, .FileName.LastIndexOf("\"))
                'End With
            Else
                frmPleaseWait1.Hide()
                Exit Sub

            End If



            'If Trim(pathIn) <> String.Empty Then
            '    If clsExcel.CreateExcelFile("ren") > 0 Then

            '        MsgBox("PrepareFileForRename() operation completed.")
            '    End If

            'End If

            If Trim(pathIn) = String.Empty Then
                Exit Sub

            End If

            Dim strList_files As List(Of String) = New List(Of String)
            Dim strList_newfiles As List(Of String) = New List(Of String)

            Dim strList_skipfiles As List(Of String) = New List(Of String)

            Dim strListList_OldfilesNewfiles As List(Of List(Of String)) = New List(Of List(Of String))()


            frmPleaseWait1.TextBox1.Text = "Selected path:" & pathIn
            frmPleaseWait1.TextBox1.Text = frmPleaseWait1.TextBox1.Text & Environment.NewLine & "Listing original files"
            frmPleaseWait1.Refresh()


            'System.Threading.Thread.Sleep(10000)

            If mode = PrepareFileForRenameMode.FilesInAllSubDir Then
                '===== List all files (no dir) in all subdir
                'strList_files.Clear()
                'strList_newfiles.Clear()
                'strListList_newfiles.Clear()


                strList_files = ProcessDirectory(pathIn)



            ElseIf mode = PrepareFileForRenameMode.FilesAndFoldersInFirstLevel Then

                '===== List all files+ dir in 1st level dir
                'strList_files.Clear()
                'strList_newfiles.Clear()
                'strListList_newfiles.Clear()


                Dim di0 As IO.DirectoryInfo = New IO.DirectoryInfo(pathIn)
                Dim fi1 As IO.FileInfo() = di0.GetFiles()
                Dim di1 As IO.DirectoryInfo() = di0.GetDirectories()

                'Dim fiTemp As IO.FileInfo


                For Each diTemp As IO.DirectoryInfo In di1
                    Dim strTemp As String

                    'strTemp = varFromJson.Item("filenames").Item("newFilenameFormat")
                    'strTemp = strTemp.Replace("{f}", Path.Combine(Path.GetDirectoryName(strFilename), intRow.ToString))
                    'strTemp = Path.Combine(di0.FullName, strTemp)
                    strTemp = diTemp.FullName


                    strList_files.Add(strTemp)
                    'intRow += 1


                Next
                For Each fiTemp As IO.FileInfo In fi1
                    Dim strTemp As String

                    'strTemp = varFromJson.Item("filenames").Item("newFilenameFormat")
                    'strTemp = strTemp.Replace("{f}", Path.Combine(Path.GetDirectoryName(strFilename), intRow.ToString))
                    'strTemp = Path.Combine(di0.FullName, strTemp)
                    strTemp = fiTemp.FullName

                    strList_files.Add(strTemp)


                Next


            End If

            frmPleaseWait1.ProgressBar1.Minimum = 0
            frmPleaseWait1.ProgressBar1.Maximum = strList_files.Count
            frmPleaseWait1.ProgressBar1.Value = 0
            frmPleaseWait1.TextBox1.Text = frmPleaseWait1.TextBox1.Text & Environment.NewLine & "Listing new files"
            frmPleaseWait1.Refresh()

            Dim intRow As Integer = 0

            Dim intLoop As Integer = 0

            For Each strFilename In strList_files.ToList
                Dim strPath = Path.GetDirectoryName(strFilename)
                Dim strName = Path.GetFileName(strFilename)


                Dim objTemp = varFromJson.Item("filenames").Item("newFilename")

                If objTemp IsNot Nothing Then

                    If Not RegExpIsMatch(strName, objTemp.Item("RegExpFindWhat")) Then

                        strList_files.Remove(strFilename)
                        strList_skipfiles.Add(strFilename)


                        intLoop += 1
                        'If intLoop > strList_files.Count * 0.02 Then

                        'frmPleaseWait1.ProgressBar1.Value = frmPleaseWait1.ProgressBar1.Value + intLoop
                        frmPleaseWait1.ProgressBar1.Value = intLoop

                        frmPleaseWait1.Refresh()
                        Continue For
                    End If


                    strName = RegExpReplace(strName, objTemp.Item("RegExpFindWhat"), objTemp.Item("RegExpReplaceWith"))

                    'strTemp = strTemp.Replace("{f}", Path.Combine(Path.GetDirectoryName(strFilename), intRow.ToString))
                    'strTemp = Path.Combine(strPath, strName)
                    'strTemp = strTemp.Replace("{f}", intRow.ToString)
                    'strList_newfiles.Add(Path.Combine(Path.GetDirectoryName(strFilename), intRow.ToString & ".doc"))

                End If



                'objTemp = varFromJson.Item("filenames").Item("newFilenameLastFormat")

                'If objTemp IsNot Nothing Then

                '    'strTemp = strTemp.Replace("{f}", Path.Combine(Path.GetDirectoryName(strFilename), intRow.ToString))
                '    'strTemp = Path.Combine(Path.GetDirectoryName(strFilename), strTemp)
                '    strName = strName.Replace("{intRow}", intRow.ToString)
                '    'strList_newfiles.Add(Path.Combine(Path.GetDirectoryName(strFilename), intRow.ToString & ".doc"))

                'End If

                strName = strName.Replace("{intRow}", intRow.ToString)

                strName = Path.Combine(strPath, strName)

                If strFilename = strName Then

                    strList_files.Remove(strFilename)
                    strList_skipfiles.Add(strFilename)

                Else
                    strList_newfiles.Add(strName)
                    intRow += 1
                End If


                intLoop += 1
                'If intLoop > strList_files.Count * 0.02 Then

                'frmPleaseWait1.ProgressBar1.Value = frmPleaseWait1.ProgressBar1.Value + intLoop
                frmPleaseWait1.ProgressBar1.Value = intLoop

                frmPleaseWait1.Refresh()
                'System.Threading.Thread.Sleep(100)

                'intLoop = 0
                'End If

            Next


            strList_files.Insert(0, "Filenames 0")
            strList_newfiles.Insert(0, "Filenames 1")



            'MsgBox(strList_files.ToArray)

            'clsExcel.listToExcel("ren", strList_files)

            'Dim strListList_newfiles As List(Of List(Of String)) = New List(Of List(Of String))()
            strListList_OldfilesNewfiles.Add(strList_files)
            strListList_OldfilesNewfiles.Add(strList_newfiles)


            frmPleaseWait1.TextBox1.Text = frmPleaseWait1.TextBox1.Text & Environment.NewLine & "Writing results to excel"
            frmPleaseWait1.Refresh()

            'clsExcel.listToExcel(StrMakeFullPath(varFromJson.Item("filenames").Item("excel4rename")), strList_files)
            clsExcel.listToExcel(StrMakeFullPath(varFromJson.Item("filenames").Item("excel4rename")), varFromJson.Item("strings").Item("worksheet.name"), strListList_OldfilesNewfiles)

            clsExcel.listToExcel(StrMakeFullPath(varFromJson.Item("filenames").Item("excel4rename")), varFromJson.Item("strings").Item("worksheet.nameForSkipFiles"), strList_skipfiles)

            frmPleaseWait1.TextBox1.Text = frmPleaseWait1.TextBox1.Text & Environment.NewLine & "Finish :D"
            frmPleaseWait1.Refresh()

            ' Hide the please wait form
            frmPleaseWait1.Hide()


            If strList_skipfiles.Count > 0 And strList_skipfiles.Count < 10 Then
                MessageBox.Show(String.Join(Environment.NewLine, strList_skipfiles), "Unmatched filename", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)


            End If

            
            MsgBox("Files have been listed in excel file." & Environment.NewLine & "strList_skipfiles.Count = " & strList_skipfiles.Count)

            ' Set cursor as default arrow
            Cursor.Current = Cursors.Default



        Catch ex As Exception
            MsgBox(GetExceptionInfo(ex))
        End Try
    End Sub


    
End Module
