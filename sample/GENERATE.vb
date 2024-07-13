Imports System.Data.SqlClient

Public Class GENERATE

    Private Sub GENERATE_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label2.Text = LOGIN.TextBox1.Text

        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"

        Dim sqlQuery As String = "SELECT c.id , c.[liable person], c.Date, c.return_date, c.reciever, " &
                         "p.item, p.quantity, p.type, p.purpose " &
                         "FROM challan c " &
                         "INNER JOIN product p ON c.id = p.ID " &  ' Assuming corrected JOIN condition
                         "ORDER BY c.id"


        Dim connection As New SqlConnection(connectionString)
        Dim adapter As New SqlDataAdapter(sqlQuery, connection)
        ' Add username parameter

        Dim dataSet As New DataSet()


        connection.Open()
            adapter.Fill(dataSet)

            If dataSet.Tables.Count > 0 Then
                DataGridView1.DataSource = dataSet.Tables(0)

        End If


            If connection.State = ConnectionState.Open Then
                connection.Close()
            End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"
        Dim id As String = TextBox1.Text

        Using connection As New SqlConnection(connectionString)
            Dim query As String = "SELECT p.item, p.quantity, p.type, p.purpose, c.[liable person], c.date, c.return_date, c.reciever " &
                                  "FROM product p " &
                                  "INNER JOIN challan c ON p.id = c.id " &
                                  "WHERE p.id = @id"
            Dim command As New SqlCommand(query, connection)
            command.Parameters.AddWithValue("@id", id)

            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()

            ' Check if any data is returned
            If reader.HasRows Then
                ' Read the common data for the challan
                reader.Read()
                Dim challanID As Integer = id
                Dim liablePerson As String = reader("liable person").ToString()
                Dim issueDate As DateTime = DateTime.Parse(reader("date").ToString())
                Dim returnDate As DateTime = DateTime.Parse(reader("return_date").ToString())
                Dim receiver As String = reader("reciever").ToString()

                ' Collect the product data
                Dim products As New List(Of (String, Integer, String, String))
                Do
                    products.Add((reader("item").ToString(),
                                  Convert.ToInt32(reader("quantity")),
                                  reader("type").ToString(),
                                  reader("purpose").ToString()))
                Loop While reader.Read()
                Me.Close()
                ' Create a new instance of SUMMARY1 form
                Dim summaryForm As New SUMMARY1()
                summaryForm.UpdateSummary(challanID, liablePerson, issueDate, returnDate, receiver, products)

                ' Open SUMMARY1 form
                summaryForm.Show()
            Else
                MessageBox.Show("No data found for the entered ID.")
            End If

            connection.Close()
        End Using
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        ADMINHOME.Show()
    End Sub
End Class