Public Class frmTextbox

    Private Sub ButtonOK_Click(sender As Object, e As EventArgs) Handles ButtonOK.Click
        strUserInput = Me.TextBox1.Text
        Me.DialogResult = DialogResult.OK
    End Sub

    Private Sub frmTextbox_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'Me.DialogResult = DialogResult.None
    End Sub

    Private Sub frmTextbox_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.TextBox1.Text = strOutput
    End Sub
End Class