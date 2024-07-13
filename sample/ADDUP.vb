Imports System.Data.SqlClient

Public Class ADDUP

    Private connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim id As Integer
        If Integer.TryParse(TextBox1.Text, id) Then
            CheckAndDisplayProducts(id)
            Panel2.Visible = True
            Panel1.Visible = False
        Else
            MessageBox.Show("Please enter a valid ID.")
            Panel2.Visible = False
        End If
    End Sub

    Private Sub CheckAndDisplayProducts(id As Integer)
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT * FROM Product WHERE id = @id ORDER BY SI"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", id)
                Dim adapter As New SqlDataAdapter(cmd)
                Dim table As New DataTable()
                adapter.Fill(table)
                If table.Rows.Count > 0 Then
                    DataGridView1.DataSource = table
                    If table.Rows.Count >= 5 Then
                        MessageBox.Show("Space full. Cannot add more items.")
                    Else
                        MessageBox.Show("Products found. You can add details.")
                    End If
                Else
                    MessageBox.Show("No products found for the given ID.")
                End If
            End Using
        End Using
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim id As Integer
        If Integer.TryParse(TextBox1.Text, id) Then
            AddProductDetail(id)
        Else
            MessageBox.Show("Please enter a valid ID.")
        End If
    End Sub

    Private Sub AddProductDetail(id As Integer)
        Dim item As String = TextBox2.Text
        Dim quantity As Integer
        Dim type As String = TextBox4.Text
        Dim purpose As String = TextBox5.Text

        If String.IsNullOrEmpty(item) OrElse Not Integer.TryParse(TextBox3.Text, quantity) OrElse String.IsNullOrEmpty(type) OrElse String.IsNullOrEmpty(purpose) Then
            MessageBox.Show("Please fill all the details correctly.")
            Return
        End If

        ' Check the current maximum SI value for the given ID
        Dim currentSI As Integer = GetCurrentMaxSI(id)
        If currentSI >= 5 Then
            MessageBox.Show("Space full. Cannot add more items.")
            Return
        End If

        ' Insert the new product detail
        Using conn As New SqlConnection(connectionString)
            conn.Open()

            Dim insertQuery As String = "INSERT INTO Product (id, item, quantity, type, purpose, SI) VALUES (@id, @item, @quantity, @type, @purpose, @SI)"
            Using insertCmd As New SqlCommand(insertQuery, conn)
                insertCmd.Parameters.AddWithValue("@id", id)
                insertCmd.Parameters.AddWithValue("@item", item)
                insertCmd.Parameters.AddWithValue("@quantity", quantity)
                insertCmd.Parameters.AddWithValue("@type", type)
                insertCmd.Parameters.AddWithValue("@purpose", purpose)
                insertCmd.Parameters.AddWithValue("@SI", currentSI + 1)
                insertCmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Product detail inserted successfully.")
            CheckAndDisplayProducts(id) ' Refresh the data grid view
        End Using
    End Sub

    Private Function GetCurrentMaxSI(id As Integer) As Integer
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim query As String = "SELECT MAX(SI) FROM Product WHERE id = @id"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", id)
                Dim result = cmd.ExecuteScalar()
                If result IsNot DBNull.Value AndAlso result IsNot Nothing Then
                    Return Convert.ToInt32(result)
                Else
                    Return 0 ' If no SI values are found, start with 0
                End If
            End Using
        End Using
    End Function

    Private Sub ADDUP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel2.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        ADMINHOME.Show()

    End Sub
End Class
