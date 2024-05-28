Imports System.IO
Imports System.Text
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing
Imports System.Diagnostics


'use 192.168.1.15 default saving
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
        TabControl1.SelectedIndex = 2
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        TabControl1.SelectedIndex = 1
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

#Region "Tab 1 Code"
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
        SaveReportToTXT()
    End Sub

    Private Sub SaveReportToTXT()
        Dim commonDirectory As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), "MIS/REPORT")
        Dim month As String = Date.Now.ToString("MMMM").ToUpper()
        Dim year As String = Date.Now.Year.ToString()
        Dim txtFileName As String = $"{month}{year}REPORT.txt"
        Dim txtFilePath As String = Path.Combine(commonDirectory, txtFileName)

        ' Save report data to the TXT file
        Dim reportData As String = $"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}" & Environment.NewLine &
                               $"Ticket Number: {TextBox1.Text}" & Environment.NewLine &
                               $"Name: {TextBox2.Text}" & Environment.NewLine &
                               $"Department: {ComboBox1.Text}" & Environment.NewLine &
                               $"Description: {RichTextBox1.Text}" & Environment.NewLine &
                               $"Level: {ComboBox2.Text}" & Environment.NewLine &
                               "-----------------------------------------" & Environment.NewLine

        Try
            If Not Directory.Exists(commonDirectory) Then
                Directory.CreateDirectory(commonDirectory)
            End If

            ' Check if the file already exists
            If File.Exists(txtFilePath) Then
                ' Append the new report data to the existing file
            End If


            Dim existingTickets As New List(Of String)

            If existingTickets.Contains(TextBox1.Text) Then
                MessageBox.Show("Error: Duplicate ticket number detected. The record was not saved.", "Duplicate Ticket Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                ' Append the new report data to the existing file or create a new file if it doesn't exist
                If File.Exists(txtFilePath) Then
                    ' Check if the ticket number already exists in the file
                    Dim fileContent As String = File.ReadAllText(txtFilePath)
                    If Not fileContent.Contains($"Ticket Number: {TextBox1.Text}") Then
                        File.AppendAllText(txtFilePath, reportData)
                    Else
                        MessageBox.Show("Error: Duplicate ticket number detected. The record was not saved.", "Duplicate Ticket Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                Else
                    File.WriteAllText(txtFilePath, reportData)
                End If

                MessageBox.Show("Report data saved to TXT file.")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error Saving TXT File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SaveReportToCSV()
        Dim commonDirectory As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), "MIS/TICKETS")
        Dim year As String = Date.Now.Year.ToString()
        Dim csvFileName As String = $"{year}REPORT.csv"
        Dim csvFilePath As String = Path.Combine(commonDirectory, csvFileName)

        ' Save report data to the CSV file
        Dim reportData As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss},{TextBox1.Text},{TextBox2.Text},{ComboBox1.Text},{RichTextBox1.Text},{ComboBox2.Text}"
        Dim newTicketNumber As String = TextBox1.Text

        Try
            If Not Directory.Exists(commonDirectory) Then
                Directory.CreateDirectory(commonDirectory)
            End If

            Dim existingTickets As New List(Of String)
            If File.Exists(csvFilePath) Then
                ' Read all existing ticket numbers from the CSV file
                Using reader As New StreamReader(csvFilePath)
                    While Not reader.EndOfStream
                        Dim line As String = reader.ReadLine()
                        Dim values As String() = line.Split(","c)
                        If values.Length > 1 Then
                            existingTickets.Add(values(1))
                        End If
                    End While
                End Using
            End If

            ' Check for duplication
            If existingTickets.Contains(newTicketNumber) Then
                MessageBox.Show("Error: Duplicate ticket number detected. The record was not saved.", "Duplicate Ticket Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                ' Write header if file does not exist
                If Not File.Exists(csvFilePath) Then
                    Dim header As String = "Date,Ticket Number,Name,Department,Description,Level"
                    File.WriteAllText(csvFilePath, header + Environment.NewLine)
                End If

                ' Append the new report data
                File.AppendAllText(csvFilePath, reportData + Environment.NewLine)
                MessageBox.Show("Report data saved to CSV file.")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error Saving CSV File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GenerateTicketNumber(currentDate As Date, departmentInitial As String) As String
        Dim ticketNumberSuffix As Integer = 1
        Dim ticketNumber As String = currentDate.ToString("ddMMyyyy") & departmentInitial & ticketNumberSuffix

        Dim commonDirectory As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), "MIS/TICKETS")
        Dim year As String = currentDate.Year.ToString()
        Dim csvFileName As String = $"{year}REPORT.csv"
        Dim csvFilePath As String = Path.Combine(commonDirectory, csvFileName)

        Dim existingTickets As New List(Of String)() ' Declare existingTickets here

        If File.Exists(csvFilePath) Then
            existingTickets = GetExistingTickets(csvFilePath, currentDate, departmentInitial)

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

        While existingTickets.Contains(currentDate.ToString("ddMMyyyy") & departmentInitial & ticketNumberSuffix.ToString())
            ticketNumberSuffix += 1
        End While

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


#End Region

#Region "Tab 2 Code"

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        'input ticket number
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        ' Trigger Find ticket
        Dim ticketNumber As String = TextBox4.Text
        Dim commonDirectory As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), "MIS/TICKETS")
        Dim year As String = Date.Now.Year.ToString()
        Dim csvFileName As String = $"{year}REPORT.csv"
        Dim csvFilePath As String = Path.Combine(commonDirectory, csvFileName)

        Try
            If Not File.Exists(csvFilePath) Then
                MessageBox.Show("CSV file does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim ticketInfo As String = FindTicketInformation(csvFilePath, ticketNumber)
            Dim departmentSection As String = ""
            Dim reportDate As String = ""
            Dim name As String = ""
            Dim description As String = ""
            Dim level As String = ""

            If Not String.IsNullOrEmpty(ticketInfo) Then
                ' Extract the fields from ticketInfo
                Dim values As String() = ticketInfo.Split(","c)
                If values.Length >= 6 Then
                    reportDate = values(0).Trim()
                    name = values(2).Trim()
                    departmentSection = values(3).Trim()
                    description = values(4).Trim()
                    level = values(5).Trim()
                Else
                    MessageBox.Show("Error: The ticket information is incomplete.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' Display information about the ticket
                Dim displayInfo As String = $"Date: {reportDate}{Environment.NewLine}Ticket Number: {ticketNumber}{Environment.NewLine}Name: {name}{Environment.NewLine}Department: {departmentSection}{Environment.NewLine}Description: {description}{Environment.NewLine}Level: {level}"
                MessageBox.Show(displayInfo, "Ticket Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' Ask user if they want to generate a PDF
                Dim dialogResult As DialogResult = MessageBox.Show("Do you want to generate a PDF for this ticket?", "Generate PDF", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If dialogResult = DialogResult.Yes Then
                    GenerateAndSavePDF(commonDirectory, ticketNumber, description, departmentSection, reportDate)
                End If
            Else
                MessageBox.Show("Ticket number not found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function FindTicketInformation(csvFilePath As String, ticketNumber As String) As String
        Using reader As New StreamReader(csvFilePath)
            While Not reader.EndOfStream
                Dim line As String = reader.ReadLine()
                Dim values As String() = line.Split(","c)
                If values.Length > 1 AndAlso values(1).Trim() = ticketNumber Then
                    Return line
                End If
            End While
        End Using
        Return String.Empty
    End Function

    Private Sub GenerateAndSavePDF(commonDirectory As String, ticketNumber As String, description As String, departmentSection As String, reportDate As String)
        Dim pdfFilePath As String = Path.Combine(commonDirectory, $"{ticketNumber}.pdf")

        ' Create a new PDF document
        Dim document As New PdfSharp.Pdf.PdfDocument()
        Dim page As PdfSharp.Pdf.PdfPage = document.AddPage()
        Dim gfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(page)
        Dim font As New PdfSharp.Drawing.XFont("Arial", 12)

        ' Draw header
        Dim italicFont As New PdfSharp.Drawing.XFont("Calibri", 10, PdfSharp.Drawing.XFontStyle.Italic)
        gfx.DrawString("F͟o͟r͟m͟ ͟R͟e͟f͟.͟Q͟F͟/͟M͟I͟S͟-͟0͟0͟1͟", italicFont, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(360, 20, 200, 0), PdfSharp.Drawing.XStringFormats.TopRight)
        gfx.DrawString("R̲e̲v̲.̲ ̲1̲;̲ ̲A̲p̲r̲i̲l̲'̲1̲8̲", italicFont, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(350, 30, 200, 0), PdfSharp.Drawing.XStringFormats.TopRight)


        Dim headerFont As New PdfSharp.Drawing.XFont("Arial", 22, PdfSharp.Drawing.XFontStyle.Bold)
        gfx.DrawString("MAYER STEEL PIPE CORPORATION", headerFont, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(0, 60, page.Width.Point, 0), PdfSharp.Drawing.XStringFormats.TopCenter)

        ' Draw MIS Department header
        Dim misHeaderFont As New PdfSharp.Drawing.XFont("Arial", 12, PdfSharp.Drawing.XFontStyle.Bold)
        gfx.DrawString("MIS Department", misHeaderFont, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(0, 85, page.Width.Point, 0), PdfSharp.Drawing.XStringFormats.TopCenter)

        ' Draw ticket number and date
        gfx.DrawString("No: ͟ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(400, 100, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString(ticketNumber, font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(430, 100, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Date: ͟ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(400, 170, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString(reportDate, font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(440, 170, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)

        ' Draw "MIS TROUBLE REPORT" header
        Dim troubleReportFont As New PdfSharp.Drawing.XFont("Arial", 14, PdfSharp.Drawing.XFontStyle.Bold)
        gfx.DrawString("MIS TROUBLE REPORT", troubleReportFont, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(0, 130, page.Width.Point, 0), PdfSharp.Drawing.XStringFormats.TopCenter)

        ' Draw department/section
        gfx.DrawString("Department/Section: ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(40, 170, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString(departmentSection, font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(180, 170, 300, 0), PdfSharp.Drawing.XStringFormats.TopLeft)

        ' Draw trouble box
        Dim troubleBoxX As Double = 40
        Dim troubleBoxY As Double = 190
        Dim troubleBoxWidth As Double = page.Width.Point * 0.4
        Dim troubleBoxHeight As Double = 135
        gfx.DrawRectangle(PdfSharp.Drawing.XPens.Black, troubleBoxX, troubleBoxY, troubleBoxWidth, troubleBoxHeight)
        gfx.DrawString("TROUBLE", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(troubleBoxX + 10, troubleBoxY + 10, troubleBoxWidth - 20, 0), PdfSharp.Drawing.XStringFormats.TopLeft)

        ' Draw the trouble description
        Dim troubleDescriptionY As Double = troubleBoxY + 30
        gfx.DrawString(description, font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(troubleBoxX + 10, troubleDescriptionY, troubleBoxWidth - 20, 0), PdfSharp.Drawing.XStringFormats.TopLeft)

        ' Draw the underline for the description
        Dim underlineHeight As Double = 0 ' Adjust the height of the underline as needed
        Dim underlineY As Double = troubleDescriptionY + font.GetHeight() - underlineHeight
        gfx.DrawRectangle(PdfSharp.Drawing.XPens.Black, troubleBoxX + 10, underlineY, troubleBoxWidth - 20, underlineHeight)

        ' Draw the additional lines
        Dim additionalLineY1 As Double = underlineY + underlineHeight + 25 ' Adjust the spacing between lines as needed
        gfx.DrawRectangle(PdfSharp.Drawing.XPens.Black, troubleBoxX + 10, additionalLineY1, troubleBoxWidth - 20, underlineHeight)

        Dim additionalLineY2 As Double = additionalLineY1 + underlineHeight + 25 ' Adjust the spacing between lines as needed
        gfx.DrawRectangle(PdfSharp.Drawing.XPens.Black, troubleBoxX + 10, additionalLineY2, troubleBoxWidth - 20, underlineHeight)

        ' Draw action box
        Dim actionBoxX As Double = page.Width.Point * 0.5
        Dim actionBoxY As Double = troubleBoxY
        Dim actionBoxWidth As Double = page.Width.Point * 0.4
        Dim actionBoxHeight As Double = 135
        gfx.DrawRectangle(PdfSharp.Drawing.XPens.Black, actionBoxX, actionBoxY, actionBoxWidth, actionBoxHeight)
        gfx.DrawString("ACTION", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(actionBoxX + 10, actionBoxY + 10, actionBoxWidth - 20, 0), PdfSharp.Drawing.XStringFormats.TopLeft)

        ' Draw action description
        Dim actionDescriptionY As Double = actionBoxY + 30
        gfx.DrawString("______________________________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(actionBoxX + 10, actionDescriptionY, actionBoxWidth - 20, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        actionDescriptionY += font.GetHeight()
        gfx.DrawString("______________________________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(actionBoxX + 10, actionDescriptionY, actionBoxWidth - 20, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        actionDescriptionY += font.GetHeight()
        gfx.DrawString("______________________________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(actionBoxX + 10, actionDescriptionY, actionBoxWidth - 20, 0), PdfSharp.Drawing.XStringFormats.TopLeft)

        ' Draw reported by section
        gfx.DrawString("Reported by:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(40, 330, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("______________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(110, 330, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Time:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(210, 330, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("______", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(240, 330, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Performed by:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(280, 330, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("______________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(355, 330, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Received by:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(40, 355, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("______________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(110, 355, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Approved by:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(40, 375, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("______________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(110, 375, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Accepted by:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(280, 355, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("____________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(350, 355, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)

        ' Draw performed by section
        gfx.DrawString("Time Started:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(450, 330, 50, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("______", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(520, 330, 150, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Time Completed:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(430, 355, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("______", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(520, 355, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)

        ' Draw the big box
        Dim bigBoxX As Double = 20
        Dim bigBoxY As Double = 160
        Dim bigBoxWidth As Double = page.Width.Point * 0.9
        Dim bigBoxHeight As Double = 260
        gfx.DrawRectangle(PdfSharp.Drawing.XPens.Black, bigBoxX, bigBoxY, bigBoxWidth, bigBoxHeight)

        ' Save the PDF document
        document.Save(pdfFilePath)

        MessageBox.Show("PDF generated and saved successfully.", "PDF Generated", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' Open the PDF file with the default viewer
        Process.Start(pdfFilePath)
    End Sub


#End Region
End Class
