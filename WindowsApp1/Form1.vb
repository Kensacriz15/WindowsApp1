Imports System.IO
Imports System.Text


Public Class Form1


    Private Sub TabPage1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TabControl1.SelectedIndex = 0
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TabControl1.SelectedIndex = 1
    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Dim form1 As New Form1()
        form1.Show()
    End Sub

    Private Function GetNextNumber() As Long
        Static LastNumber As Long = 99999
        LastNumber += 1
        Return LastNumber
    End Function

    Private Function GetNextNumber(xlWorkSheet As Object) As Long
        Dim lastRow As Integer = xlWorkSheet.UsedRange.Rows.Count

        Dim highestTicketNumber As Long = 99999
        For i As Integer = 2 To lastRow
            Dim ticketNumber As Long
            If Long.TryParse(xlWorkSheet.Cells(i, 2).Value, ticketNumber) Then
                If ticketNumber > highestTicketNumber Then
                    highestTicketNumber = ticketNumber
                End If
            End If
        Next

        Return highestTicketNumber + 1
    End Function

    Private Sub Button_Click4(sender As Object, e As EventArgs) Handles Button4.Click
        Dim xlApp As Object = CreateObject("Excel.Application")
        Dim xlWorkBook As Object = Nothing
        Dim xlWorkSheet As Object = Nothing
        Dim excelFilePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\report.xlsx"

        Try
            If File.Exists(excelFilePath) Then
                xlWorkBook = xlApp.Workbooks.Open(excelFilePath)
                xlWorkSheet = xlWorkBook.Sheets(1)

                Dim lastRow As Integer = xlWorkSheet.UsedRange.Rows.Count

                Dim ticketNumbers As New List(Of String)()
                For i As Integer = 2 To lastRow
                    ticketNumbers.Add(xlWorkSheet.Cells(i, 2).Value)
                Next

                Dim nextTicketNumber As String = GenerateNextTicketNumber(ticketNumbers, xlWorkSheet)

                Dim currentRow As Integer = lastRow + 1
                xlWorkSheet.Cells(currentRow, 1) = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                xlWorkSheet.Cells(currentRow, 2) = nextTicketNumber
                xlWorkSheet.Cells(currentRow, 3) = TextBox2.Text
                xlWorkSheet.Cells(currentRow, 4) = ComboBox1.Text
                xlWorkSheet.Cells(currentRow, 5) = RichTextBox1.Text

                TextBox1.Text = nextTicketNumber
            Else
                MessageBox.Show("Excel file not found.")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error Reading Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If xlWorkBook IsNot Nothing Then xlWorkBook.Close(False)
            ReleaseObject(xlWorkSheet)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlApp)
        End Try
    End Sub

    Private Function GenerateNextTicketNumber(ticketNumbers As List(Of String), xlWorkSheet As Object) As String
        Dim newNumber As Long = GetNextNumber(xlWorkSheet)
        Dim nextTicketNumber As String = newNumber.ToString()

        While ticketNumbers.Contains(nextTicketNumber)
            newNumber = GetNextNumber(xlWorkSheet)
            nextTicketNumber = newNumber.ToString()
        End While

        Return nextTicketNumber
    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'Generate Ticket Number
        TextBox1.ReadOnly = True
    End Sub

    Private originalText As String

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        'Name
        If String.IsNullOrEmpty(originalText) Then
            originalText = TextBox2.Text
        End If
    End Sub

    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave

        If TextBox2.Text = "" Then
            TextBox2.Text = originalText
        End If
    End Sub

    Private ReadOnly defaultSelection As Boolean = True

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ' Select Department
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'Selectiong of level
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        'Description
    End Sub

    Private Sub SaveReportToTextFile()
        Dim filePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\test.txt"
        Dim data As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "|" & TextBox1.Text & "|" & TextBox2.Text & "|" & ComboBox1.Text & "|" & ComboBox2.Text & "|" & RichTextBox1.Text

        If File.Exists(filePath) Then
            File.AppendAllText(filePath, Environment.NewLine & data)
        Else
            File.WriteAllText(filePath, data)
        End If
    End Sub

    Private Function GenerateUniqueTicketNumber(xlApp As Object, excelFilePath As String) As String
        Dim ticketNumber As String
        Dim numbersFromFile As New List(Of String)()

        If File.Exists(excelFilePath) Then
            Dim xlWorkBook As Object = xlApp.Workbooks.Open(excelFilePath)
            Dim xlWorkSheet As Object = xlWorkBook.Sheets(1)

            ' Get existing ticket numbers
            Dim lastRow As Integer = xlWorkSheet.UsedRange.Rows.Count
            For i As Integer = 2 To lastRow ' Start from row 2 (skipping headers)
                numbersFromFile.Add(xlWorkSheet.Cells(i, 2).Value)
            Next

            xlWorkBook.Close(False)
        End If

        ' Generate unique ticket number
        Dim newNumber As Long
        Do
            newNumber = GetNextNumber()
            ticketNumber = newNumber.ToString()
        Loop While numbersFromFile.Contains(ticketNumber)

        Return ticketNumber
    End Function

    Private Sub RichTextBox2_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox2.TextChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0

        Dim excelFilePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\report.xlsx"
        Dim excelContent As String = ReadExcelFile(excelFilePath)
        RichTextBox2.Text = excelContent
        Dim timer As New Timer()
        timer.Interval = 20000
        AddHandler timer.Tick, AddressOf RefreshExcelContent
        timer.Start()

        Timer1.Start()


    End Sub

    Private Function ReadExcelFile(excelFilePath As String) As String
        Dim xlApp As Object = CreateObject("Excel.Application")
        Dim xlWorkBook As Object = Nothing
        Dim xlWorkSheet As Object = Nothing
        Dim excelContent As New StringBuilder()

        Try
            If File.Exists(excelFilePath) Then
                xlWorkBook = xlApp.Workbooks.Open(excelFilePath)
                xlWorkSheet = xlWorkBook.Sheets(1)

                Dim lastRow As Integer = xlWorkSheet.UsedRange.Rows.Count

                For i As Integer = 1 To lastRow
                    Dim rowContent As New StringBuilder()
                    For j As Integer = 1 To xlWorkSheet.UsedRange.Columns.Count
                        Dim cellValue As String = xlWorkSheet.Cells(i, j).Value
                        rowContent.Append(cellValue & "   ") ' Adjust the separator as needed
                    Next
                    excelContent.AppendLine(rowContent.ToString())
                Next
            Else
                MessageBox.Show("Excel file not found.")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error Reading Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If xlWorkBook IsNot Nothing Then xlWorkBook.Close(False)
            ReleaseObject(xlWorkSheet)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlApp)
        End Try

        Return excelContent.ToString()
    End Function

    Private Sub RefreshExcelContent(sender As Object, e As EventArgs)
        Dim excelFilePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\report.xlsx"
        Dim excelContent As String = ReadExcelFile(excelFilePath)
        RichTextBox2.Text = excelContent
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Dim currentDateAndTime As DateTime = DateTime.Now
        Dim weekdayName As String = currentDateAndTime.DayOfWeek.ToString()
        TextBox3.Text = weekdayName & "  " & currentDateAndTime.ToString("MMMM dd yyyy HH:mm:ss")
        TextBox3.ReadOnly = True
        TextBox3.Enabled = False
    End Sub



    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim currentDateAndTime As DateTime = DateTime.Now
        Dim monthName As String = currentDateAndTime.ToString("MMMM")
        Dim weekdayName As String = currentDateAndTime.DayOfWeek.ToString()

        TextBox3.Text = currentDateAndTime.ToString("yyyy/MM/dd HH:mm:ss") &
                        Environment.NewLine & "Weekday: " & weekdayName &
                        Environment.NewLine & "Month: " & monthName
    End Sub

    Private Sub RichTextBox3_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub WriteToExcel()
        Dim xlApp As Object = CreateObject("Excel.Application")
        Dim xlWorkBook As Object = Nothing
        Dim xlWorkSheet As Object = Nothing

        Dim excelFilePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\report.xlsx"
        Dim ticketNumber As String = GenerateUniqueTicketNumber(xlApp, excelFilePath)

        Try
            If File.Exists(excelFilePath) Then
                xlWorkBook = xlApp.Workbooks.Open(excelFilePath)
                xlWorkSheet = xlWorkBook.Sheets(1)

                Dim lastRow As Integer = xlWorkSheet.UsedRange.Rows.Count

                ' Check if the header row is missing and write headers if needed
                If lastRow = 1 Then
                    xlWorkSheet.Cells(1, 1) = "Date"
                    xlWorkSheet.Cells(1, 2) = "Ticket Number"
                    xlWorkSheet.Cells(1, 3) = "Name"
                    xlWorkSheet.Cells(1, 4) = "Department"
                    xlWorkSheet.Cells(1, 5) = "Description"
                    lastRow += 1
                End If

                ' Write new data to the next row
                Dim currentRow As Integer = lastRow + 1
                xlWorkSheet.Cells(currentRow, 1) = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                xlWorkSheet.Cells(currentRow, 2) = ticketNumber
                xlWorkSheet.Cells(currentRow, 3) = TextBox2.Text
                xlWorkSheet.Cells(currentRow, 4) = ComboBox1.Text
                xlWorkSheet.Cells(currentRow, 5) = RichTextBox1.Text

                ' Save Excel file
                xlWorkBook.Save()
                File.SetAttributes(excelFilePath, FileAttributes.Normal) ' Set file attributes to normal
                MessageBox.Show("Report data saved to Excel file.")
            Else
                MessageBox.Show("Excel file not found.")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error Saving Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If xlWorkBook IsNot Nothing Then xlWorkBook.Close(True)
            ReleaseObject(xlWorkSheet)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlApp)
        End Try
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        SaveReportToTextFile()
        WriteToExcel()
        MessageBox.Show("Report Ticket Submitted")
    End Sub

End Class