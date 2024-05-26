Imports System.IO
Imports System.Text
Imports iText.Kernel.Exceptions
Imports iText.Kernel.Pdf
Imports iText.Layout
Imports iText.Layout.Element
Imports Microsoft.VisualBasic.Logging

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        Timer1.Start()
    End Sub
    Private Sub TabPage1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TabControl1.SelectedIndex = 0
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TabControl1.SelectedIndex = 1
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        TabControl1.SelectedIndex = 2
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim currentDateAndTime As DateTime = DateTime.Now
        Dim monthName As String = currentDateAndTime.ToString("MMMM")
        Dim weekdayName As String = currentDateAndTime.DayOfWeek.ToString()

        TextBox3.Text = currentDateAndTime.ToString("yyyy/MM/dd HH:mm:ss") &
                        Environment.NewLine & "Weekday: " & weekdayName &
                        Environment.NewLine & "Month: " & monthName
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Dim currentDateAndTime As DateTime = DateTime.Now
        Dim weekdayName As String = currentDateAndTime.DayOfWeek.ToString()
        TextBox3.Text = weekdayName & "  " & currentDateAndTime.ToString("MMMM dd yyyy HH:mm:ss")
        TextBox3.ReadOnly = True
        TextBox3.Enabled = False
    End Sub

    'TAB 1
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ' Select Department CSMC, MKK, SPIRAL, MARKETING, LOGISTIC
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        ' Name
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Trigger Ticket Number Generation
        Dim currentDate As Date = Date.Now
        Dim departmentInitial As String = ComboBox1.SelectedItem.ToString().Substring(0, 2)
        Dim ticketNumber As String = GenerateTicketNumber(currentDate, departmentInitial)
        TextBox1.Text = ticketNumber
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs)
        ' Description Problem
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ' Level LOW, URGENT, CRITICAL
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Save Report Data to CSV File
        SaveReportToCSV()
    End Sub

    Private Sub SaveReportToCSV()
        Dim commonDirectory As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), "MIS/REPORT")
        Dim year As String = Date.Now.Year.ToString()
        Dim csvFileName As String = $"{year}REPORT.csv"
        Dim csvFilePath As String = Path.Combine(commonDirectory, csvFileName)

        ' Save report data to the CSV file
        Dim reportData As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss},{TextBox1.Text},{TextBox2.Text},{ComboBox1.Text},{RichTextBox1.Text},{ComboBox2.Text}"

        Try
            If Not Directory.Exists(commonDirectory) Then
                Directory.CreateDirectory(commonDirectory)
            End If

            If Not File.Exists(csvFilePath) Then
                Dim header As String = "Date,Ticket Number,Name,Department,Description,Level"
                File.WriteAllText(csvFilePath, header + Environment.NewLine)
            End If

            File.AppendAllText(csvFilePath, reportData + Environment.NewLine)

            MessageBox.Show("Report data saved to CSV file.")
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error Saving CSV File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GenerateTicketNumber(currentDate As Date, departmentInitial As String) As String
        Dim ticketNumberSuffix As Integer = 1
        Dim ticketNumber As String = currentDate.ToString("ddMMyyyy") & departmentInitial & ticketNumberSuffix

        Dim commonDirectory As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), "MIS/REPORT")
        Dim year As String = currentDate.Year.ToString()
        Dim csvFileName As String = $"{year}REPORT.csv"
        Dim csvFilePath As String = Path.Combine(commonDirectory, csvFileName)

        If File.Exists(csvFilePath) Then
            Dim existingTickets As List(Of String) = GetExistingTickets(csvFilePath, currentDate, departmentInitial)

            If existingTickets.Count > 0 Then
                Dim validTicketNumbers = New List(Of Integer)
                For Each ticket In existingTickets
                    Dim suffixStr As String = ticket.Substring(8)
                    Dim suffix As Integer
                    If Integer.TryParse(suffixStr.Substring(2), suffix) Then
                        validTicketNumbers.Add(suffix)
                    End If
                Next

                If validTicketNumbers.Count > 0 Then
                    Dim maxSuffix = validTicketNumbers.Max()
                    ticketNumberSuffix = maxSuffix + 1
                End If
            End If
        End If

        ticketNumber = currentDate.ToString("ddMMyyyy") & departmentInitial & ticketNumberSuffix.ToString()
        Return ticketNumber
    End Function

    Private Function GetExistingTickets(csvFilePath As String, currentDate As Date, departmentInitial As String) As List(Of String)
        Dim existingTickets As New List(Of String)()

        Using reader As New StreamReader(csvFilePath)
            While Not reader.EndOfStream
                Dim line As String = reader.ReadLine()
                Dim values As String() = line.Split(","c)
                If values.Length > 1 Then
                    Dim datePart As String = currentDate.ToString("ddMMyyyy")
                    If values(1).StartsWith(datePart & departmentInitial) Then
                        existingTickets.Add(values(1))
                    End If
                End If
            End While
        End Using

        Return existingTickets
    End Function




End Class
