Public Class frmPleaseWait

    Private Sub frmPleaseWait_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True

        Me.Hide()
    End Sub

    Private Sub frmPleaseWait_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class