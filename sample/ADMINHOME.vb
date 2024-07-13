
Imports System.Data.SqlClient

Public Class ADMINHOME
    Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"
    Private Sub LOGOUTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LOGOUTToolStripMenuItem.Click
        Dim result As DialogResult
        result = MessageBox.Show("ARE YOU SURE YOU WANT TO RESET THE DATABASE", "Confirmation",
        MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        If result = DialogResult.OK Then
            Dim tables As String() = {"CHALLAN", "PRODUCT", "STATUS"}

            Using connection As New SqlConnection(connectionString)
                connection.Open()

                For Each table As String In tables
                    Dim deleteCommand As New SqlCommand($"DELETE FROM {table}", connection)
                    deleteCommand.ExecuteNonQuery()
                Next

                MessageBox.Show("Data deleted from all tables successfully.")
            End Using
        End If
    End Sub
    Private Sub GENERATEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GENERATEToolStripMenuItem.Click
        Me.Close()
        GENERATE.Show()
    End Sub

    Private Sub CHECKSTATUSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CHECKSTATUSToolStripMenuItem.Click
        Me.Close()
        STATUS.Show()
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DELETEToolStripMenuItem.Click
        Me.Close()
        DEL.Show()
    End Sub

    Private Sub POSTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles POSTToolStripMenuItem.Click
        Me.Close()
        POST.Show()
    End Sub

    Private Sub PRODUCTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PRODUCTToolStripMenuItem.Click
        Me.Close()
        ADD.Show()
    End Sub
    Private Sub LIABLEPERSONToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LIABLEPERSONToolStripMenuItem.Click
        Me.Close()
        LIABLE.Show()
    End Sub

    Private Sub PARTYToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PARTYToolStripMenuItem.Click
        Me.Close()
        PARTY.Show()
    End Sub
    Private Sub PENDINNGToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PENDINNGToolStripMenuItem.Click
        Me.Close()
        PENDING.Show()
    End Sub
    Private Sub EDITToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EDITToolStripMenuItem.Click
        Me.Close()
        UPDATE1.Show()
    End Sub
    Private Sub ADDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDToolStripMenuItem.Click
        Me.Close()
        ADDUP.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim result As DialogResult
        result = MessageBox.Show("ARE YOU SURE YOU WANT TO LOGOUT", "Confirmation",
        MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        If result = DialogResult.OK Then
            Me.Close()
            Form1.Show()
        End If
    End Sub
End Class