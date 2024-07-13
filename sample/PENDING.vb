Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Data
Imports System.Drawing
Imports System.IO

Public Class PENDING
    Private Sub PENDING_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=return;Integrated Security=True;Encrypt=False"

        ' SQL query to fetch data from PRODUCT and CHALLAN tables based on STATUS table's conditions
        Dim query As String = "
           SELECT  
            C.id AS ChallanID, C.[liable person], C.[reciever], C.date,C.RETURN_DATE, P.item, P.quantity,P.TYPE,P.PURPOSE
        FROM PRODUCT P
        JOIN CHALLAN C ON P.ID = C.id -- Ensure this is the correct join condition
        WHERE NOT EXISTS (
            SELECT 1 FROM STATUS S WHERE S.id = P.id AND S.SI = P.SI
        )"

        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    connection.Open()
                    Dim adapter As New SqlDataAdapter(command)
                    Dim dataTable As New System.Data.DataTable()
                    adapter.Fill(dataTable)

                    ' Set DataGridView properties
                    DataGridView1.DataSource = dataTable

                    ' Auto-resize columns
                    DataGridView1.AutoResizeColumns()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"An error occurred: {ex.Message}")
        End Try
    End Sub

    Public Sub ExportDataGridViewToExcel()
        Dim excelApp As New Application()
        Dim excelWorkbook As Workbook = Nothing
        Dim excelWorksheet As Worksheet = Nothing

        Try
            excelWorkbook = excelApp.Workbooks.Add()
            excelWorksheet = excelWorkbook.Sheets(1)
            excelWorksheet.Name = "Pending Challan"

            ' Merge and center-align cells for header
            Dim headerRange As Range = excelWorksheet.Range("A1:H1")
            headerRange.Merge()
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter
            headerRange.Value = "Motilal Dulichand"
            headerRange.Font.Size = 36
            headerRange.Font.Bold = True

            ' Merge and center-align cells for title
            Dim titleRange As Range = excelWorksheet.Range("A2:H2")
            titleRange.Merge()
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter
            titleRange.Value = "Pending Challan"
            titleRange.Font.Size = 24
            titleRange.Font.Bold = True

            ' Adding DataGridView headers horizontally
            For A As Integer = 1 To DataGridView1.Columns.Count
                excelWorksheet.Cells(4, A).Value = DataGridView1.Columns(A - 1).HeaderText
                excelWorksheet.Cells(4, A).Font.Bold = True
                excelWorksheet.Cells(4, A).Interior.Color = RGB(195, 240, 247)
                excelWorksheet.Cells(4, A).Font.Color = RGB(0, 0, 0)
            Next A

            For I As Integer = 0 To DataGridView1.Rows.Count - 1
                For j As Integer = 0 To DataGridView1.Columns.Count - 1
                    Dim columnName As String = DataGridView1.Columns(j).Name
                    If columnName = "date" Or columnName = "RETURN_DATE" Then
                        If DataGridView1.Rows(I).Cells(j).Value IsNot Nothing AndAlso Not String.IsNullOrEmpty(DataGridView1.Rows(I).Cells(j).Value.ToString()) Then
                            Try
                                Dim dateValue As Date = Date.Parse(DataGridView1.Rows(I).Cells(j).Value.ToString())
                                excelWorksheet.Cells(5 + I, 1 + j).Value = dateValue.ToString("dd-MM-yyyy")
                            Catch ex As Exception
                                ' Handle parsing errors, if necessary
                                excelWorksheet.Cells(5 + I, 1 + j).Value = DataGridView1.Rows(I).Cells(j).Value.ToString()
                            End Try
                        End If
                    Else
                        If DataGridView1.Rows(I).Cells(j).Value IsNot Nothing Then
                            excelWorksheet.Cells(5 + I, 1 + j).Value = DataGridView1.Rows(I).Cells(j).Value.ToString()
                        End If
                    End If
                Next
            Next

            ' Auto-fit the columns
            excelWorksheet.Columns.AutoFit()

            ' Open SaveFileDialog to ask for the save location
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Files|*.xlsx"
            saveFileDialog.Title = "Save an Excel File"
            saveFileDialog.FileName = $"PendingChallan_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.xlsx"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                ' Save the Excel file to the selected path
                Dim filePath As String = saveFileDialog.FileName
                excelWorkbook.SaveAs(filePath)

                ' Notify the user
                MessageBox.Show($"File saved to: {filePath}")
            End If


        Catch ex As Exception
            MessageBox.Show($"An error occurred while exporting to Excel and printing: {ex.Message}")
        Finally
            ' Release resources
            If excelWorksheet IsNot Nothing Then releaseObject(excelWorksheet)
            If excelWorkbook IsNot Nothing Then
                excelWorkbook.Close(False)
                releaseObject(excelWorkbook)
            End If
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                releaseObject(excelApp)
            End If
        End Try
    End Sub

    Public Sub PrintExcelFile(filePath As String)
        Dim excelApp As New Application()
        Dim excelWorkbook As Workbook = Nothing

        Try
            excelWorkbook = excelApp.Workbooks.Open(filePath)

            ' Set print settings
            With excelWorkbook.ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
                .CenterHorizontally = False
                .CenterVertically = False
                .Orientation = XlPageOrientation.xlLandscape
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .TopMargin = excelApp.InchesToPoints(0.75)

                .LeftMargin = excelApp.InchesToPoints(0.5)
                .RightMargin = excelApp.InchesToPoints(0.5)
                .BottomMargin = excelApp.InchesToPoints(0.5)
            End With

            ' Print the Excel file
            excelWorkbook.PrintOut()

            ' Close the workbook without saving changes
            excelWorkbook.Close(False)
        Catch ex As Exception
            MessageBox.Show($"An error occurred while printing: {ex.Message}")
        Finally
            ' Release resources
            If excelWorkbook IsNot Nothing Then releaseObject(excelWorkbook)
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                releaseObject(excelApp)
            End If
        End Try
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ExportDataGridViewToExcel()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        ADMINHOME.Show()
    End Sub
End Class
