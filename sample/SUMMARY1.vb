Imports System.Data.SqlClient
Imports System.Drawing.Printing

Public Class SUMMARY1

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        ADMINHOME.Show()
    End Sub
    Public Sub New()
        ' Initialize form components
        InitializeComponent()
    End Sub

    ' Method to update the form with provided data
    Public Sub UpdateSummary(challanID As Integer, liablePerson As String, issueDate As DateTime, returnDate As DateTime, receiver As String, products As List(Of (String, Integer, String, String)))
        Label3.Text = challanID.ToString()
        Label9.Text = liablePerson
        Label11.Text = issueDate.ToString("dd MMM yyyy")
        Label43.Text = returnDate.ToString("dd MMM yyyy")
        Label7.Text = receiver
        UpdateAddress()
        ' Update product labels and set visibility
        If products.Count > 0 Then
            Label17.Visible = True
            Label22.Text = products(0).Item1
            Label22.Visible = True
            Label23.Text = products(0).Item2.ToString()
            Label23.Visible = True
            Label24.Text = products(0).Item3
            Label24.Visible = True
            Label25.Text = products(0).Item4
            Label25.Visible = True
        Else
            Label22.Visible = False
            Label23.Visible = False
            Label24.Visible = False
            Label25.Visible = False
            Label17.Visible = False
        End If

        If products.Count > 1 Then
            Label18.Visible = True
            Label26.Text = products(1).Item1
            Label26.Visible = True
            Label27.Text = products(1).Item2.ToString()
            Label27.Visible = True
            Label28.Text = products(1).Item3
            Label28.Visible = True
            Label29.Text = products(1).Item4
            Label29.Visible = True
        Else
            Label26.Visible = False
            Label27.Visible = False
            Label28.Visible = False
            Label29.Visible = False
            Label18.Visible = False
        End If

        If products.Count > 2 Then
            Label19.Visible = True
            Label30.Text = products(2).Item1
            Label30.Visible = True
            Label31.Text = products(2).Item2.ToString()
            Label31.Visible = True
            Label32.Text = products(2).Item3
            Label32.Visible = True
            Label33.Text = products(2).Item4
            Label33.Visible = True
        Else
            Label30.Visible = False
            Label31.Visible = False
            Label32.Visible = False
            Label33.Visible = False
            Label19.Visible = False
        End If

        If products.Count > 3 Then
            Label20.Visible = True
            Label34.Text = products(3).Item1
            Label34.Visible = True
            Label35.Text = products(3).Item2.ToString()
            Label35.Visible = True
            Label36.Text = products(3).Item3
            Label36.Visible = True
            Label37.Text = products(3).Item4
            Label37.Visible = True
        Else
            Label34.Visible = False
            Label35.Visible = False
            Label36.Visible = False
            Label37.Visible = False
            Label20.Visible = False
        End If

        If products.Count > 4 Then
            Label21.Visible = True
            Label38.Text = products(4).Item1
            Label38.Visible = True
            Label39.Text = products(4).Item2.ToString()
            Label39.Visible = True
            Label40.Text = products(4).Item3
            Label40.Visible = True
            Label41.Text = products(4).Item4
            Label41.Visible = True
        Else
            Label38.Visible = False
            Label39.Visible = False
            Label40.Visible = False
            Label41.Visible = False
            Label21.Visible = False
        End If
    End Sub
    Private Sub PrintDocument_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim font As New Font("Arial", 10)
        Dim headerFont As New Font("Arial", 12, FontStyle.Bold)
        Dim linePen As New Pen(Color.Black, 1)
        Dim headerFont1 As New Font("Arial", 24, FontStyle.Bold, 1.5)
        Dim font1 As New Font("Arial", 10, FontStyle.Bold)

        ' Set margins
        Dim leftMargin As Integer = e.MarginBounds.Left
        Dim topMargin As Integer = e.MarginBounds.Top
        Dim rightMargin As Integer = e.MarginBounds.Right

        ' Company Information
        e.Graphics.DrawString("RETURNABLE CHALLAN", headerFont, Brushes.Black, leftMargin + 440, topMargin)
        e.Graphics.DrawString("Motilal  Dulichand  (P)   Ltd.", headerFont1, Brushes.Black, leftMargin + 390, topMargin + 20)
        e.Graphics.DrawString("20, Industrial Estate, Kanpur-208012", font1, Brushes.Gray, leftMargin + 440, topMargin + 50)
        e.Graphics.DrawString("GSTIN NO: 09AAAACD8231ZZ", font, Brushes.Black, leftMargin, topMargin)

        ' Challan Information
        e.Graphics.DrawString($"Challan No: {Label3.Text}", font, Brushes.Black, leftMargin, topMargin + 80)
        e.Graphics.DrawString($"Challan Dt: {Label11.Text}", font, Brushes.Black, leftMargin + 760, topMargin + 80)

        ' Recipient Information
        e.Graphics.DrawString("TO,", font, Brushes.Black, leftMargin, topMargin + 120)
        e.Graphics.DrawString(Label7.Text, font, Brushes.Black, leftMargin, topMargin + 140)
        e.Graphics.DrawString(Label44.Text, font, Brushes.Black, leftMargin, topMargin + 160)
        e.Graphics.DrawString($"Liable Person: {Label9.Text}", font, Brushes.Black, leftMargin + 400, topMargin + 160)

        ' Table Header
        Dim yPos As Integer = topMargin + 200
        e.Graphics.DrawLine(linePen, leftMargin, yPos, rightMargin, yPos)

        e.Graphics.DrawString("S. No.", font, Brushes.Black, leftMargin, yPos)
        e.Graphics.DrawString("PARTICULARS", font, Brushes.Black, leftMargin + 50, yPos)
        e.Graphics.DrawString("Quantity", font, Brushes.Black, leftMargin + 500, yPos)
        e.Graphics.DrawString("TYPE", font, Brushes.Black, leftMargin + 610, yPos)
        e.Graphics.DrawString("Purpose", font, Brushes.Black, leftMargin + 700, yPos)

        ' Draw lines for the table header
        yPos += 20
        e.Graphics.DrawLine(linePen, leftMargin, yPos, rightMargin, yPos)

        ' Table Data
        If Label22.Visible = True Then
            yPos += 20
            e.Graphics.DrawString("1", font, Brushes.Black, leftMargin + 10, yPos)
            e.Graphics.DrawString(Label22.Text, font, Brushes.Black, leftMargin + 50, yPos)
            e.Graphics.DrawString(Label23.Text, font, Brushes.Black, leftMargin + 500, yPos)
            e.Graphics.DrawString(Label24.Text, font, Brushes.Black, leftMargin + 610, yPos)
            e.Graphics.DrawString(Label25.Text, font, Brushes.Black, leftMargin + 700, yPos)
        End If

        If Label26.Visible = True Then
            yPos += 30
            e.Graphics.DrawString("2", font, Brushes.Black, leftMargin + 10, yPos)
            e.Graphics.DrawString(Label26.Text, font, Brushes.Black, leftMargin + 50, yPos)
            e.Graphics.DrawString(Label27.Text, font, Brushes.Black, leftMargin + 500, yPos)
            e.Graphics.DrawString(Label28.Text, font, Brushes.Black, leftMargin + 610, yPos)
            e.Graphics.DrawString(Label29.Text, font, Brushes.Black, leftMargin + 700, yPos)
        End If
        If Label30.Visible = True Then
            yPos += 30
            e.Graphics.DrawString("3", font, Brushes.Black, leftMargin + 10, yPos)
            e.Graphics.DrawString(Label30.Text, font, Brushes.Black, leftMargin + 50, yPos)
            e.Graphics.DrawString(Label31.Text, font, Brushes.Black, leftMargin + 500, yPos)
            e.Graphics.DrawString(Label32.Text, font, Brushes.Black, leftMargin + 610, yPos)
            e.Graphics.DrawString(Label33.Text, font, Brushes.Black, leftMargin + 700, yPos)
        End If
        If Label34.Visible = True Then
            yPos += 30
            e.Graphics.DrawString("4", font, Brushes.Black, leftMargin + 10, yPos)
            e.Graphics.DrawString(Label34.Text, font, Brushes.Black, leftMargin + 50, yPos)
            e.Graphics.DrawString(Label35.Text, font, Brushes.Black, leftMargin + 500, yPos)
            e.Graphics.DrawString(Label36.Text, font, Brushes.Black, leftMargin + 610, yPos)
            e.Graphics.DrawString(Label37.Text, font, Brushes.Black, leftMargin + 700, yPos)
        End If
        If Label38.Visible = True Then
            yPos += 30
            e.Graphics.DrawString("5", font, Brushes.Black, leftMargin + 10, yPos)
            e.Graphics.DrawString(Label38.Text, font, Brushes.Black, leftMargin + 50, yPos)
            e.Graphics.DrawString(Label39.Text, font, Brushes.Black, leftMargin + 500, yPos)
            e.Graphics.DrawString(Label40.Text, font, Brushes.Black, leftMargin + 610, yPos)
            e.Graphics.DrawString(Label41.Text, font, Brushes.Black, leftMargin + 700, yPos)
        End If
        yPos += 20
        ' Draw lines for the table data
        yPos += 40
        e.Graphics.DrawLine(linePen, leftMargin, yPos, rightMargin, yPos)

        ' Draw vertical lines to form the table
        Dim columnPositions As Integer() = {leftMargin, leftMargin + 40, leftMargin + 490, leftMargin + 590, leftMargin + 690, rightMargin}
        For Each xPos As Integer In columnPositions
            e.Graphics.DrawLine(linePen, xPos, topMargin + 200, xPos, yPos)
        Next

        ' Footer Information
        yPos += 20
        e.Graphics.DrawString($"Date of Return: {Label43.Text}", font, Brushes.Black, leftMargin, yPos)
        e.Graphics.DrawString($"Through: PARTY", font, Brushes.Black, leftMargin + 700, yPos)
        yPos += 20
        e.Graphics.DrawLine(linePen, leftMargin, yPos, rightMargin, yPos)

        yPos += 20
        e.Graphics.DrawString("Sign. Mech Store Incharge", font, Brushes.Black, leftMargin, yPos)
        e.Graphics.DrawString("Forwarded By (Section Head)", font, Brushes.Black, leftMargin + 350, yPos)
        e.Graphics.DrawString("Approved By", font, Brushes.Black, leftMargin + 700, yPos)

        ' Draw lines for the footer
        yPos += 20
        e.Graphics.DrawLine(linePen, leftMargin, yPos, rightMargin, yPos)
        yPos += 20
        e.Graphics.DrawString("material returned in store", font, Brushes.Black, leftMargin, yPos)
        e.Graphics.DrawString("quality checked", font, Brushes.Black, leftMargin + 350, yPos)
        e.Graphics.DrawString("posted computer", font, Brushes.Black, leftMargin + 700, yPos)
        yPos += 15
        e.Graphics.DrawString("sign mech store incharge", font, Brushes.Black, leftMargin, yPos)
        e.Graphics.DrawString("by mech Section Head)", font, Brushes.Black, leftMargin + 350, yPos)
        e.Graphics.DrawString("dt.        by.", font, Brushes.Black, leftMargin + 700, yPos)
        yPos += 30
        e.Graphics.DrawLine(linePen, leftMargin, yPos, rightMargin, yPos)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PrintDocument1.DefaultPageSettings.Landscape = True
        PrintDocument1.Print()
    End Sub
    Public Sub UpdateAddress()
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"
        Dim query As String = "SELECT Address FROM Party WHERE [RECIEVER] = @Receiver"

        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@Receiver", Label7.Text)

                Try
                    connection.Open()
                    Dim address As Object = command.ExecuteScalar()

                    If address IsNot Nothing Then
                        Label44.Text = address.ToString()
                    Else
                        Label44.Text = "Address not found."
                    End If
                Catch ex As Exception
                    MessageBox.Show("An error occurred: " & ex.Message)
                End Try
            End Using
        End Using

    End Sub
    Private Sub SomeMethod()
        ' Set Label7.Text value here
        Label7.Text = "Receiver Name"

        ' Call the method to update Label46
        UpdateAddress()
    End Sub
End Class