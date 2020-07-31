Public Class aboutForm
    Private Sub aboutForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtText.Text = vbCrLf + vbCrLf + vbCrLf
        txtText.Text += "   NETD - 2202" + vbCrLf + vbCrLf
        txtText.Text += "   LAB #5" + vbCrLf + vbCrLf
        txtText.Text += "   Hoang Quoc Bao Nguyen"
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        Me.Close()
    End Sub
End Class