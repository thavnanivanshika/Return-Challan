Imports System.Data.SqlClient

Public Class LIABLE

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim liable As String = TextBox1.Text.Trim()
        Dim address As String = TextBox2.Text.Trim()
        Dim city As String = TextBox3.Text.Trim()
        Dim pincode As String = TextBox4.Text.Trim()
        Dim phone As String = TextBox5.Text.Trim()

        ' Check if any field is empty
        If String.IsNullOrEmpty(liable) OrElse
           String.IsNullOrEmpty(address) OrElse
           String.IsNullOrEmpty(city) OrElse
           String.IsNullOrEmpty(pincode) OrElse
           String.IsNullOrEmpty(phone) Then

            MessageBox.Show("Please fill in all fields.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Validate phone number format
        If Not System.Text.RegularExpressions.Regex.IsMatch(phone, "^\d{10}$") Then
            MessageBox.Show("Phone number should be exactly ten digits.", "Invalid Phone Number", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"

        ' SQL query to insert data into the LIABLE table
        Dim query As String = "INSERT INTO [dbo].[LIABLE] (LIABLE, ADDRESS, CITY, PINCODE, PHONE) " &
                              "VALUES (@LIABLE, @Address, @City, @Pincode, @Phone)"

        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    ' Add parameters to prevent SQL injection
                    command.Parameters.AddWithValue("@LIABLE", liable)
                    command.Parameters.AddWithValue("@Address", address)
                    command.Parameters.AddWithValue("@City", city)
                    command.Parameters.AddWithValue("@Pincode", pincode)
                    command.Parameters.AddWithValue("@Phone", phone)

                    connection.Open()
                    command.ExecuteNonQuery()

                    MessageBox.Show("Data inserted successfully.")
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"An error occurred: {ex.Message}")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        ADMINHOME.Show()
    End Sub

End Class
