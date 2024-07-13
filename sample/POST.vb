Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class POST
    Dim CONNECTION_STRING As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        ADMINHOME.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PostSelectedRowToStatus()
    End Sub

    Private Sub PostSelectedRowToStatus()
        ' Check if a row is selected
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Get values from the selected row
            Dim id As String = selectedRow.Cells("id").Value.ToString()
            Dim si As String = selectedRow.Cells("SI").Value.ToString()

            ' Determine the value for the status based on the selected radio button
            Dim status As String
            If RadioButton1.Checked Then
                status = "YES"
            ElseIf RadioButton2.Checked Then
                status = "NO"
            Else
                MessageBox.Show("Please select a status (Yes or No).")
                Return
            End If

            ' Example SQL query to insert into STATUS table
            Dim insertQuery As String = "INSERT INTO STATUS (ID, POST, SI) VALUES (@ID, @POST, @SI)"

            Try
                Using connection As New SqlConnection(CONNECTION_STRING)
                    connection.Open()

                    ' Insert data into the STATUS table
                    Using insertCommand As New SqlCommand(insertQuery, connection)
                        ' Parameters to prevent SQL injection
                        insertCommand.Parameters.AddWithValue("@ID", id)
                        insertCommand.Parameters.AddWithValue("@POST", status)
                        insertCommand.Parameters.AddWithValue("@SI", si)

                        insertCommand.ExecuteNonQuery()
                        MessageBox.Show("Data inserted into STATUS table successfully.")

                        ' Remove the posted row from DataGridView1
                        DataGridView1.Rows.Remove(selectedRow)
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show($"An error occurred: {ex.Message}")
            End Try
        Else
            MessageBox.Show("Please select a row to post to STATUS.")
        End If
    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"

        ' Get the value from TextBox1 (assuming it's the ID you want to search for)
        Dim valueD As String = TextBox1.Text.Trim()

        ' SQL query to select data from PRODUCT table based on ID
        Dim selectProductQuery As String = "SELECT id, item, quantity, type, purpose, SI FROM PRODUCT WHERE id = @ID"

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                ' Retrieve data from PRODUCT table for the selected ID
                Using selectProductCommand As New SqlCommand(selectProductQuery, connection)
                    selectProductCommand.Parameters.AddWithValue("@ID", valueD)

                    ' Create a DataTable to hold the results
                    Dim dataTable As New DataTable()
                    Dim adapter As New SqlDataAdapter(selectProductCommand)
                    adapter.Fill(dataTable)

                    ' Bind the DataTable to DataGridView1
                    DataGridView1.DataSource = dataTable
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"An error occurred: {ex.Message}")
        End Try
    End Sub
End Class
