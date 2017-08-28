Imports System.IO


Public Class frmBase

    Protected numberOfButtons As Integer
    Protected buttons() As Button



    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AddButtons()

    End Sub

    Sub CheckNONExist()

        strOutput = strUserInput

        Dim frm As New frmTextbox
        If frm.ShowDialog = DialogResult.OK Then

            'Dim strFilesNonExist As String = ""

            Dim lstFilesNonExist As List(Of String) = New List(Of String)


            strUserInput = strUserInput
            For Each strFilename In strUserInput.Replace(vbLf, "").Split(Environment.NewLine)
                If Not File.Exists(strFilename) Then
                    'strFilesNonExist &= strFilename & Environment.NewLine
                    lstFilesNonExist.Add(strFilename)
                End If
            Next

            strOutput = String.Format("There are {0} NON-Existing files{1}{2}",                                      lstFilesNonExist.Count,
                                      Environment.NewLine,
                                      String.Join(Environment.NewLine, lstFilesNonExist.ToArray()))

            Dim frm2 As New frmTextbox
            If frm2.ShowDialog = DialogResult.OK Then

            End If
        End If


    End Sub

    Protected Sub AddButtons()
        'numberOfButtons = 3
        numberOfButtons = varFromJson.Item("buttons").Count

        ReDim buttons(numberOfButtons)
        'For counter As Integer = 0 To numberOfButtons - 1

        Dim counter As Integer = 0

        For Each btn In varFromJson.Item("buttons")

            'For counter As Integer = 0 To varFromJson.Item("excelDependent").Count



            buttons(counter) = New Button
            With buttons(counter)


                .Size = New Drawing.Size(100, 20)
                .Visible = True
                .Location = New Drawing.Size(55, 33 + counter * 30)
                '.Text = "Button " + (counter + 1).ToString ' or some name from an array you pass from main
                .Text = btn.Item("text")
                .AutoSize = True
                .AutoSizeMode = Windows.Forms.AutoSizeMode.GrowOnly

                'any other property

                AddHandler buttons(counter).Click, AddressOf All_Buttons_Clicked

                Me.Controls.Add(buttons(counter))
                'MsgBox(.Text)
            End With
            '

            counter = counter + 1
        Next
    End Sub

    Protected Sub All_Buttons_Clicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'some code here, can check to see which checkbox was changed, which button was clicked, by number or text

        Dim btn As Button = DirectCast(sender, Button)

        For Each btnJson In varFromJson.Item("buttons")
            If btn.Text = btnJson.Item("text") Then
                CallByName(Me, btnJson.Item("function"), CallType.Method)

            End If
        Next


        'If btn.Text.ToUpper = "BUTTON 1" Then

        '    'Dim txtResult As String
        '    If frmDatagrid.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
        '        ' Read the contents of testDialog's TextBox.
        '        'txtResult = frmDatagrid.TextBox1.Text
        '    Else
        '        'txtResult = "Cancelled"
        '    End If
        '    frmDatagrid.Dispose()
        'ElseIf btn.Text.ToUpper = "BUTTON 2" Then

        '    MsgBox(varFromJson.Item("excelFormat").Item("address").Item("heading"))


        'End If

    End Sub


    Public Sub PrepareFileForRenameFirstLevel()
        PrepareFileForRename(PrepareFileForRenameMode.FilesAndFoldersInFirstLevel)
    End Sub

    Public Sub PrepareFileForRenameAllSubDir()
        PrepareFileForRename(PrepareFileForRenameMode.FilesInAllSubDir)
    End Sub

    Sub RenameFileBasedOnExcel1()
        RenameFileBasedOnExcel("ren")
    End Sub

    Sub RenameFileBasedOnExcel2()
        RenameFileBasedOnExcel("renForFistLevel")
    End Sub

    Sub OpenFileExcel()
        System.Diagnostics.Process.Start(varFromJson.Item("filenames").Item("excel4rename"))
    End Sub

    Sub OpenFileJson()
        System.Diagnostics.Process.Start(strJsonPath)
    End Sub

End Class
