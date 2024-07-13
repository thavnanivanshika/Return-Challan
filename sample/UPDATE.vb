Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class UPDATE1
    Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"
    Private Sub UPDATE_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadDataGridView() ' Load data into DataGridView on form load
        ListBox1.Visible = False
        ListBox2.Visible = False
    End Sub

    ' Method to load data into DataGridView
    Private Sub LoadDataGridView()
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"
        Dim sqlQuery As String = "SELECT c.id, c.[liable person], c.Date, c.return_date, c.reciever, " &
                                 "p.si,p.item, p.quantity, p.type, p.purpose " &
                                 "FROM challan c " &
                                 "INNER JOIN product p ON c.id = p.ID " &
                                 "ORDER BY c.id"

        Dim connection As New SqlConnection(connectionString)
        Dim adapter As New SqlDataAdapter(sqlQuery, connection)
        Dim dataSet As New DataSet()

        connection.Open()
        adapter.Fill(dataSet)

        If dataSet.Tables.Count > 0 Then
            DataGridView1.DataSource = dataSet.Tables(0)
        End If

        connection.Close()
    End Sub

    ' Event handler for row selection in DataGridView
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            TextBox1.Text = selectedRow.Cells("id").Value.ToString()
            TextBox8.Text = selectedRow.Cells("si").Value.ToString() ' Assuming you have a hidden TextBox for the ID
            TextBox2.Text = selectedRow.Cells("item").Value.ToString()
            TextBox3.Text = selectedRow.Cells("quantity").Value.ToString()
            TextBox4.Text = selectedRow.Cells("type").Value.ToString()
            TextBox5.Text = selectedRow.Cells("purpose").Value.ToString()
            TextBox6.Text = selectedRow.Cells("liable person").Value.ToString()
            DateTimePicker1.Value = DateTime.Parse(selectedRow.Cells("date").Value.ToString())
            DateTimePicker2.Value = DateTime.Parse(selectedRow.Cells("return_date").Value.ToString())
            TextBox7.Text = selectedRow.Cells("reciever").Value.ToString()
        End If
    End Sub

    ' Button2_Click event handler to update the data
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a row to update.")
            Return
        End If
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"
        Dim id As String = TextBox1.Text ' Use the hidden TextBox to get the ID
        Dim si As String = TextBox8.Text
        Using connection As New SqlConnection(connectionString)
            Dim query As String = "UPDATE product SET item = @item, quantity = @quantity, type = @type, purpose = @purpose WHERE id = @id AND si = @si;" &
                                  "UPDATE challan SET [liable person] = @liablePerson, date = @date, return_date = @returnDate, reciever = @reciever WHERE id = @id;"
            Dim command As New SqlCommand(query, connection)
            command.Parameters.AddWithValue("@id", id)
            command.Parameters.AddWithValue("@si", si)
            command.Parameters.AddWithValue("@item", TextBox2.Text)
            command.Parameters.AddWithValue("@quantity", Convert.ToInt32(TextBox3.Text))
            command.Parameters.AddWithValue("@type", TextBox4.Text)
            command.Parameters.AddWithValue("@purpose", TextBox5.Text)
            command.Parameters.AddWithValue("@liablePerson", TextBox6.Text)
            command.Parameters.AddWithValue("@date", DateTimePicker1.Value)
            command.Parameters.AddWithValue("@returnDate", DateTimePicker2.Value)
            command.Parameters.AddWithValue("@reciever", TextBox7.Text)

            connection.Open()
            Dim rowsAffected As Integer = command.ExecuteNonQuery()

            If rowsAffected > 0 Then
                MessageBox.Show("Data updated successfully.")
                LoadDataGridView() ' Reload data in DataGridView
            Else
                MessageBox.Show("Update failed.")
            End If

            connection.Close()
        End Using
    End Sub
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs)
        ListBox1.Visible = False
        FetchLiablePersons(TextBox6.Text.Trim())
    End Sub

    Private Sub FetchLiablePersons(partialName As String)
        ListBox1.Visible = True
        ' Clear existing items in ListBox1
        ListBox1.Items.Clear()

        If String.IsNullOrEmpty(partialName) Then
            Return
        End If

        Dim query As String = "SELECT LIABLE FROM LIABLE WHERE LIABLE LIKE @PartialName"
        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    ' Add parameter for partial name search
                    command.Parameters.AddWithValue("@PartialName", partialName + "%")

                    connection.Open()

                    Using reader As SqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            ListBox1.Items.Add(reader("LIABLE").ToString())
                        End While
                    End Using

                    ' Check if no names were found
                    If ListBox1.Items.Count = 0 Then
                        ListBox1.Items.Add("No liable person found.")
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"An error occurred: {ex.Message}")
        End Try
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ' Set selected item in TextBox1
        If ListBox1.SelectedIndex <> -1 Then
            TextBox6.Text = ListBox1.SelectedItem.ToString()

        End If
        ListBox1.Visible = False
    End Sub
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs)
        ListBox2.Visible = True
        FetchLiablePersons1(TextBox7.Text.Trim())

    End Sub

    Private Sub FetchLiablePersons1(partialName As String)
        ListBox2.Visible = True
        ' Clear existing items in ListBox1
        ListBox2.Items.Clear()

        If String.IsNullOrEmpty(partialName) Then
            Return
        End If

        Dim query As String = "SELECT RECIEVER FROM PARTY WHERE RECIEVER LIKE @PartialName"
        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    ' Add parameter for partial name search
                    command.Parameters.AddWithValue("@PartialName", partialName + "%")

                    connection.Open()

                    Using reader As SqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            ListBox2.Items.Add(reader("RECIEVER").ToString())
                        End While
                    End Using

                    ' Check if no names were found
                    If ListBox2.Items.Count = 0 Then
                        ListBox2.Items.Add("No PARTY found.")
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"An error occurred: {ex.Message}")
        End Try
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        ' Set selected item in TextBox1
        If ListBox2.SelectedIndex <> -1 Then
            TextBox7.Text = ListBox2.SelectedItem.ToString()

        End If
        ListBox2.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        ADMINHOME.Show()
    End Sub
End Class
