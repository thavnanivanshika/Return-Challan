Public Class LOGIN
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim username As String = TextBox1.Text.Trim()
        Dim password As String = TextBox2.Text
        If String.IsNullOrWhiteSpace(username) OrElse String.IsNullOrWhiteSpace(password) Then
            MessageBox.Show("Username and password are required.")
            Return
        End If
        If username = "VANSHIKA" AndAlso password = "12345" Then
            Me.Hide()
            ADMINHOME.Show()
        Else
            MessageBox.Show("Invalid username or password.")
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox2.UseSystemPasswordChar = False
        ElseIf CheckBox1.Checked = False Then
            TextBox2.UseSystemPasswordChar = True
        End If
    End Sub
    Private Sub LOGIN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If CheckBox1.Checked = False Then
            TextBox2.UseSystemPasswordChar = True
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Form1.Show()
    End Sub

End Class
