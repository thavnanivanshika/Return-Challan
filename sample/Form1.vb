Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        LOGIN.Show()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit the application?",
"Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        ' If the user confirms, exit the application
        If result = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub
End Class
