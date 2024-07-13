Imports System.Data.SqlClient

Public Class ADD
    Dim CONNECTION_STRING As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"
    Private Sub ADD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.Visible = False
        ListBox2.Visible = False
        GroupBox1.Visible = False
        GroupBox2.Visible = False
        GroupBox3.Visible = False
        GroupBox4.Visible = False
        GroupBox5.Visible = False
        Button1.Enabled = False
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem = "1" Then
            GroupBox1.Visible = True
            GroupBox2.Visible = False
            GroupBox3.Visible = False
            GroupBox4.Visible = False
            GroupBox5.Visible = False
        ElseIf ComboBox1.SelectedItem = "2" Then
            GroupBox1.Visible = True
            GroupBox2.Visible = True
            GroupBox3.Visible = False
            GroupBox4.Visible = False
            GroupBox5.Visible = False
        ElseIf ComboBox1.SelectedItem = "3" Then
            GroupBox1.Visible = True
            GroupBox2.Visible = True
            GroupBox3.Visible = True
            GroupBox4.Visible = False
            GroupBox5.Visible = False
        ElseIf ComboBox1.SelectedItem = "4" Then
            GroupBox1.Visible = True
            GroupBox2.Visible = True
            GroupBox3.Visible = True
            GroupBox4.Visible = True
            GroupBox5.Visible = False
        ElseIf ComboBox1.SelectedItem = "5" Then
            GroupBox1.Visible = True
            GroupBox2.Visible = True
            GroupBox3.Visible = True
            GroupBox5.Visible = True
            GroupBox4.Visible = True
        Else
            MessageBox.Show("PLEASE SELECT ")
        End If
        ValidateForm()
    End Sub
    Private Sub ValidateForm()
        ' Check if all required textboxes are filled
        Dim allFilled As Boolean = Not (String.IsNullOrWhiteSpace(TextBox22.Text) Or
                                    String.IsNullOrWhiteSpace(TextBox21.Text) Or
                                    (GroupBox1.Visible AndAlso (String.IsNullOrWhiteSpace(TextBox1.Text) Or String.IsNullOrWhiteSpace(TextBox2.Text) Or String.IsNullOrWhiteSpace(TextBox3.Text) Or String.IsNullOrWhiteSpace(TextBox4.Text))) Or
                                    (GroupBox2.Visible AndAlso (String.IsNullOrWhiteSpace(TextBox8.Text) Or String.IsNullOrWhiteSpace(TextBox7.Text) Or String.IsNullOrWhiteSpace(TextBox6.Text) Or String.IsNullOrWhiteSpace(TextBox5.Text))) Or
                                    (GroupBox3.Visible AndAlso (String.IsNullOrWhiteSpace(TextBox12.Text) Or String.IsNullOrWhiteSpace(TextBox11.Text) Or String.IsNullOrWhiteSpace(TextBox10.Text) Or String.IsNullOrWhiteSpace(TextBox9.Text))) Or
                                    (GroupBox4.Visible AndAlso (String.IsNullOrWhiteSpace(TextBox16.Text) Or String.IsNullOrWhiteSpace(TextBox15.Text) Or String.IsNullOrWhiteSpace(TextBox14.Text) Or String.IsNullOrWhiteSpace(TextBox13.Text))) Or
                                    (GroupBox5.Visible AndAlso (String.IsNullOrWhiteSpace(TextBox20.Text) Or String.IsNullOrWhiteSpace(TextBox19.Text) Or String.IsNullOrWhiteSpace(TextBox18.Text) Or String.IsNullOrWhiteSpace(TextBox17.Text))))

        ' Enable or disable the button based on whether all required fields are filled
        Button1.Enabled = allFilled
    End Sub
    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged, TextBox21.TextChanged, TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged, TextBox8.TextChanged, TextBox7.TextChanged, TextBox6.TextChanged, TextBox5.TextChanged, TextBox12.TextChanged, TextBox11.TextChanged, TextBox10.TextChanged, TextBox9.TextChanged, TextBox16.TextChanged, TextBox15.TextChanged, TextBox14.TextChanged, TextBox13.TextChanged, TextBox20.TextChanged, TextBox19.TextChanged, TextBox18.TextChanged, TextBox17.TextChanged
        ValidateForm()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Using connection As New SqlConnection(CONNECTION_STRING)
            Try
                connection.Open()
                ' Generate unique ID
                Dim newID As Integer = GenerateUniqueID(connection)
                ' Insert into challan table
                Dim liablePerson As String = TextBox22.Text
                Dim issueDate As DateTime = DateTimePicker1.Value
                Dim returnDate As DateTime = DateTimePicker2.Value
                Dim receiver As String = TextBox21.Text
                InsertChallan(connection, newID, liablePerson, issueDate, returnDate, receiver)
                ' Insert into product table
                ' Insert into product table
                Dim products As List(Of (String, Integer, String, String, Integer)) = GetProducts()
                For Each product In products
                    InsertProduct(connection, newID, product.Item1, product.Item2, product.Item3, product.Item4, product.Item5)
                Next
                Me.Close()
                Dim COMBO As String = ComboBox2.SelectedItem.ToString()
                ' Show the summary form
                Dim summaryForm As New Summary()
                summaryForm.UpdateSummary(newID, liablePerson, issueDate, returnDate, receiver, products, COMBO)
                summaryForm.ShowDialog()
                MessageBox.Show("Records inserted successfully.", "Success")
            Catch ex As Exception
                MessageBox.Show("An error occurred: " & ex.Message, "Error")
            Finally
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                End If
            End Try
        End Using
    End Sub
    Private Function GenerateUniqueID(connection As SqlConnection) As Integer
        Dim newID As Integer
        ' Define the query to get the highest ID from the table
        Dim query As String = "SELECT ISNULL(MAX([id]), 9999) FROM [challan]"
        Using command As New SqlCommand(query, connection)
            ' Execute the query and get the highest ID
            Dim hiGhestID As Integer = Convert.ToInt32(command.ExecuteScalar())
            ' Increment the highest ID by one for the new ID
            newID = highestID + 1
        End Using
        Return newID
    End Function
    Private Sub InsertChallan(connection As SqlConnection, id As Integer, liablePerson As String, issueDate As DateTime, returnDate As DateTime, receiver As String)
        Dim query As String = "INSERT INTO [challan] ([id], [liable person], [date], [return_date], [reciever]) VALUES (@id, @liablePerson, @date, @returnDate, @receiver)"
        Using command As New SqlCommand(query, connection)
            command.Parameters.AddWithValue("@id", id)
            command.Parameters.AddWithValue("@liablePerson", liablePerson)
            command.Parameters.AddWithValue("@date", issueDate)
            command.Parameters.AddWithValue("@returnDate", returnDate)
            command.Parameters.AddWithValue("@receiver", receiver)
            command.ExecuteNonQuery()
        End Using
    End Sub
    Private Sub InsertProducts(connection As SqlConnection, id As Integer)
        Dim products As New List(Of (String, Integer, String, String, Integer))
        ' Read data from GroupBox1
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
            products.Add((TextBox1.Text, Convert.ToInt32(TextBox2.Text), TextBox3.Text, TextBox4.Text, 1))
        End If
        ' Read data from GroupBox2
        If Not String.IsNullOrWhiteSpace(TextBox8.Text) Then
            products.Add((TextBox8.Text, Convert.ToInt32(TextBox7.Text), TextBox6.Text, TextBox5.Text, 2))
        End If
        ' Read data from GroupBox3
        If Not String.IsNullOrWhiteSpace(TextBox12.Text) Then
            products.Add((TextBox12.Text, Convert.ToInt32(TextBox11.Text), TextBox10.Text, TextBox9.Text, 3))
        End If
        ' Read data from GroupBox4
        If Not String.IsNullOrWhiteSpace(TextBox16.Text) Then
            products.Add((TextBox16.Text, Convert.ToInt32(TextBox15.Text), TextBox14.Text, TextBox13.Text, 4))
        End If
        ' Read data from GroupBox5
        If Not String.IsNullOrWhiteSpace(TextBox20.Text) Then
            products.Add((TextBox20.Text, Convert.ToInt32(TextBox19.Text), TextBox18.Text, TextBox17.Text, 5))
        End If
        ' Insert each product into the product table
        For Each product In products
            InsertProduct(connection, id, product.Item1, product.Item2, product.Item3, product.Item4, product.Item5)
        Next
    End Sub
    Private Sub InsertProduct(connection As SqlConnection, id As Integer, item As String, quantity As Integer, type As String, purpose As String, si As Integer)
        Dim query As String = "INSERT INTO [product] ([id], [item], [quantity], [type], [purpose], [si]) VALUES (@id, @item, @quantity, @type, @purpose, @si)"
        Using command As New SqlCommand(query, connection)
            command.Parameters.AddWithValue("@id", id)
            command.Parameters.AddWithValue("@item", item)
            command.Parameters.AddWithValue("@quantity", quantity)
            command.Parameters.AddWithValue("@type", type)
            command.Parameters.AddWithValue("@purpose", purpose)
            command.Parameters.AddWithValue("@si", si)
            command.ExecuteNonQuery()
        End Using
    End Sub
    Private Function GetProducts() As List(Of (String, Integer, String, String, Integer))
        Dim products As New List(Of (String, Integer, String, String, Integer))

        ' Read data from GroupBox1
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
            products.Add((TextBox1.Text, Convert.ToInt32(TextBox2.Text), TextBox3.Text, TextBox4.Text, 1))
        End If

        ' Read data from GroupBox2
        If Not String.IsNullOrWhiteSpace(TextBox8.Text) Then
            products.Add((TextBox8.Text, Convert.ToInt32(TextBox7.Text), TextBox6.Text, TextBox5.Text, 2))
        End If

        ' Read data from GroupBox3
        If Not String.IsNullOrWhiteSpace(TextBox12.Text) Then
            products.Add((TextBox12.Text, Convert.ToInt32(TextBox11.Text), TextBox10.Text, TextBox9.Text, 3))
        End If

        ' Read data from GroupBox4
        If Not String.IsNullOrWhiteSpace(TextBox16.Text) Then
            products.Add((TextBox16.Text, Convert.ToInt32(TextBox15.Text), TextBox14.Text, TextBox13.Text, 4))
        End If

        ' Read data from GroupBox5
        If Not String.IsNullOrWhiteSpace(TextBox20.Text) Then
            products.Add((TextBox20.Text, Convert.ToInt32(TextBox19.Text), TextBox18.Text, TextBox17.Text, 5))
        End If

        Return products
    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        ADMINHOME.Show()
    End Sub
    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged
        ListBox1.Visible = False
        FetchLiablePersons(TextBox22.Text.Trim())
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
            Using connection As New SqlConnection(CONNECTION_STRING)
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
            TextBox22.Text = ListBox1.SelectedItem.ToString()

        End If
        ListBox1.Visible = False
    End Sub
    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged
        ListBox2.Visible = False
        FetchLiablePersons1(TextBox21.Text.Trim())
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
            Using connection As New SqlConnection(CONNECTION_STRING)
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
            TextBox21.Text = ListBox2.SelectedItem.ToString()
        End If
        ListBox2.Visible = False
    End Sub
End Class