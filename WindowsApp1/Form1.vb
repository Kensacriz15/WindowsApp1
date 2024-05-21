Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel


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

    Private UsedNumbers As New List(Of Long)()
    Private NumbersFromFile As New List(Of Long)()

    Private Sub ReadNumbersFromFile()
        Dim filePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\test.txt"
        If File.Exists(filePath) Then
            Dim lines As String() = File.ReadAllLines(filePath)
            For Each line As String In lines
                Dim number As Long
                If Long.TryParse(line, number) Then
                    NumbersFromFile.Add(number)
                End If
            Next
        End If
    End Sub

    Public Function GenerateUnique5DigitNumber() As Long
        Dim newNumber As Long
        Do
            newNumber = GetNextNumber()
        Loop While UsedNumbers.Contains(newNumber) Or NumbersFromFile.Contains(newNumber)
        If Not NumbersFromFile.Contains(newNumber) Then
            UsedNumbers.Add(newNumber)
        End If
        Return newNumber
    End Function

    Private Function GetNextNumber() As Long
        Static LastNumber As Long = 100000
        LastNumber += 1
        Return LastNumber
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ReadNumbersFromFile()
        Dim newNumber As Long = GenerateUnique5DigitNumber()
        TextBox1.Text = newNumber.ToString()
    End Sub

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

    Private defaultSelection As Boolean = True

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'Select Department
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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        SaveReportToTextFile()
        WriteToExcel()
        MessageBox.Show("Report Ticket Submitted")
    End Sub

    Private Sub WriteToExcel()
        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add()
        Dim xlWorkSheet As Excel.Worksheet = CType(xlWorkBook.Sheets(1), Excel.Worksheet)

        Try
            ' Write headers
            xlWorkSheet.Cells(1, 1) = "Date"
            xlWorkSheet.Cells(1, 2) = "Ticket Number"
            xlWorkSheet.Cells(1, 3) = "Name"
            xlWorkSheet.Cells(1, 4) = "Department"
            xlWorkSheet.Cells(1, 5) = "Description"

            ' Write data
            Dim row As Integer = 2
            xlWorkSheet.Cells(row, 1) = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            xlWorkSheet.Cells(row, 2) = TextBox1.Text
            xlWorkSheet.Cells(row, 3) = TextBox2.Text

            ' Save Excel file
            Dim excelFilePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\report.xlsx"
            xlWorkBook.SaveAs(excelFilePath)
            MessageBox.Show("Report data saved to Excel file.")
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error Saving Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ReleaseObject(xlWorkSheet)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlApp)
        End Try
    End Sub

    Private Sub RichTextBox2_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox2.TextChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim filePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\test.txt"
        Dim fileContent As String = File.ReadAllText(filePath)
        RichTextBox2.Text = fileContent

        Dim timer As New Timer()
        timer.Interval = 2000
        AddHandler timer.Tick, AddressOf RefreshFileContent
        timer.Start()

        Timer1.Start()
    End Sub


    Private Sub RefreshFileContent(sender As Object, e As EventArgs)
        Dim filePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\test.txt"
        Dim fileContent As String = File.ReadAllText(filePath)
        RichTextBox2.Text = fileContent
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

End Class