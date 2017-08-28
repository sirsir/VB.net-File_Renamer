Imports System.Text.RegularExpressions

Module modRegExp

    Public Function RegExpReplace(ByVal strIn As String, ByVal regexpFindWhat As String, ByVal regexpReplaceWith As String) As String


        If regexpFindWhat = "" Then
            Return strIn
        End If

        Dim regexpTemp As Regex = New Regex(regexpFindWhat)
        'MsgBox(regexp)
        'MsgBox(strIn)

        If regexpTemp.IsMatch(strIn) Then
            'MessageBox.Show(message, title, button, icon)
            'MyMessageBox()
            Return regexpTemp.Replace(strIn, regexpReplaceWith)
        Else
            Return strIn
        End If


    End Function

    Public Function RegExpIsMatch(ByVal strIn As String, ByVal regexp As String) As Boolean


        If regexp = "" Then
            Return True
        End If

        Dim regexpTemp As Regex = New Regex(regexp)
        'MsgBox(regexp)
        'MsgBox(strIn)

        If Not regexpTemp.IsMatch(strIn) Then
            'MessageBox.Show(message, title, button, icon)
            'MyMessageBox()
            Return False
        Else
            Return True
        End If


    End Function
End Module
