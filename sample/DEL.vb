Imports System.Data.SqlClient

Public Class DEL
    Private Sub DEL_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"

        Dim sqlQuery As String = "Select c.id, c.[liable person], c.Date, c.return_date, c.reciever, p.item, p.quantity, p.type, p.purpose ,P.SI 
From challan c  
INNER Join product p ON c.id = p.ID 
ORDER BY c.id;"





        Dim connection As New SqlConnection(connectionString)
        Dim adapter As New SqlDataAdapter(sqlQuery, connection)
        ' Add username parameter

        Dim dataSet As New DataSet()


        connection.Open()
        adapter.Fill(dataSet)

        If dataSet.Tables.Count > 0 Then
            DataGridView2.DataSource = dataSet.Tables(0)

        End If


        If connection.State = ConnectionState.Open Then
            connection.Close()
        End If

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView2.Rows.Count > 0 Then
            ' Get the index of the last row
            Dim lastRowIndex As Integer = DataGridView2.Rows.Count - 2 ' -2 because of the new row at the e
            ' Get the index of the currently selected row
            Dim selectedRowIndex As Integer = DataGridView2.CurrentCell.RowIndex

            ' Check if the selected row is the last row
            If selectedRowIndex = lastRowIndex Then
                ' Get the ID of the last row
                Dim lastRowId As Integer = Convert.ToInt32(DataGridView2.Rows(lastRowIndex).Cells("id").Value)
                Dim lastRowId1 As Integer = Convert.ToInt32(DataGridView2.Rows(lastRowIndex).Cells("SI").Value)

                ' Delete the last row from the database
                Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"
                Using connection As New SqlConnection(connectionString)
                    connection.Open()
                    Dim deleteQuery As String = "DELETE p 
                             FROM PRODUCT p
                             INNER JOIN CHALLAN c ON p.id = c.id
                             WHERE p.id = @id AND p.SI = @SI;

                            DELETE c
                             FROM CHALLAN c
                             WHERE c.id = @id
                             AND NOT EXISTS (SELECT 1 FROM PRODUCT p WHERE p.id = c.id AND p.SI = 1);"

                    Using command As New SqlCommand(deleteQuery, connection)
                        command.Parameters.AddWithValue("@id", lastRowId)
                        command.Parameters.AddWithValue("@SI", lastRowId1)
                        command.ExecuteNonQuery()
                    End Using
                End Using

                ' Delete the last row from the DataGridView
                DataGridView2.Rows.RemoveAt(lastRowIndex)
            Else
                ' Show a message box if the selected row is not the last row
                MessageBox.Show("You can only delete the last row.", "Delete Row", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Else
            ' Show a message box if there are no rows to delete
            MessageBox.Show("There are no rows to delete.", "Delete Row", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        ADMINHOME.Show()
    End Sub
End Class
