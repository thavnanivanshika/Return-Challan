Imports System.Data.SqlClient

Public Class STATUS
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        ADMINHOME.Show()
    End Sub

    Private Sub STATUS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Visible = False
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Get the entered Challan ID
        Dim challanID As String = TextBox1.Text.Trim()

        ' Check if the Challan ID is not empty
        If String.IsNullOrEmpty(challanID) Then
            MessageBox.Show("Please enter a Challan ID")
            Return
        End If

        ' Connection string to your database
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"

        ' SQL Query to check Challan ID and get details
        Dim query As String = "SELECT  c.id AS challan_id, c.[liable person], c.[date], c.[return_date], c.[reciever], p.[item], p.[quantity], p.[type], p.[purpose]
                               FROM challan c
                               LEFT JOIN product p ON c.id = p.id
                               WHERE c.id = @challan_id"

        ' Using block to ensure the connection is closed and disposed properly
        Using connection As New SqlConnection(connectionString)
            Try
                ' Open the connection
                connection.Open()

                ' Create the command
                Using command As New SqlCommand(query, connection)
                    ' Add parameter to the query
                    command.Parameters.AddWithValue("@challan_id", challanID)

                    ' Create a data adapter
                    Using adapter As New SqlDataAdapter(command)
                        ' Create a data table to hold the results
                        Dim dataTable As New DataTable()

                        ' Fill the data table
                        adapter.Fill(dataTable)

                        ' Check if any rows were returned
                        If dataTable.Rows.Count > 0 Then

                            ' Show the status in a message box
                            MessageBox.Show("Status: Active")
                            DataGridView1.Visible = True
                            ' Bind the data table to the DataGridView
                            DataGridView1.DataSource = dataTable
                        Else
                            DataGridView1.Visible = False
                            ' Show a message if the Challan ID does not exist
                            MessageBox.Show("Status: Inactive. Challan ID does not exist.")
                        End If
                    End Using
                End Using
            Catch ex As Exception
                ' Show an error message if something goes wrong
                MessageBox.Show("An error occurred: " & ex.Message)
            End Try
        End Using
    End Sub
End Class
