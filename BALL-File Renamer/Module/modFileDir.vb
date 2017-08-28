Imports System
Imports System.IO
Imports System.Collections


Module modFileDir
    Function ProcessDirectory(ByVal targetDirectory As String) As List(Of String)
        Dim strListReturn As List(Of String)
        strListReturn = New List(Of String)

        Dim fileEntries As String() = Directory.GetFiles(targetDirectory)
        ' Process the list of files found in the directory.
        Dim fileName As String
        For Each fileName In fileEntries
            ProcessFile(fileName)
            strListReturn.Add(fileName)

        Next fileName
        Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
        ' Recurse into subdirectories of this directory.
        Dim subdirectory As String
        For Each subdirectory In subdirectoryEntries
            strListReturn.AddRange(ProcessDirectory(subdirectory))
        Next subdirectory

        Return strListReturn

    End Function 'ProcessDirectory

    ' Insert logic for processing found files here.
    Sub ProcessFile(ByVal path As String)
        Console.WriteLine("Processed file '{0}'.", path)
    End Sub 'ProcessFile


    Function StrMakeFullPath(strPathIn) As String
        Return Path.Combine(Environment.CurrentDirectory, strPathIn.ToString)

    End Function
End Module
