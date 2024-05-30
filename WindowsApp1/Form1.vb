Imports System.IO
Imports System.Text
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Security.Principal

'use 192.168.1.15 default saving
Public Class Form1

    Public Class Impersonation
        <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Private Shared Function LogonUser(lpszUsername As String, lpszDomain As String, lpszPassword As String, dwLogonType As Integer, dwLogonProvider As Integer, ByRef phToken As IntPtr) As Boolean
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function CloseHandle(handle As IntPtr) As Boolean
        End Function

        Private Const LOGON32_LOGON_INTERACTIVE As Integer = 2
        Private Const LOGON32_PROVIDER_DEFAULT As Integer = 0
        Private impersonationContext As WindowsImpersonationContext

        Public Function ImpersonateValidUser(userName As String, domain As String, password As String) As Boolean
            Dim tokenHandle As New IntPtr(0)

            If LogonUser(userName, domain, password, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, tokenHandle) <> 0 Then
                Dim newId As New WindowsIdentity(tokenHandle)
                impersonationContext = newId.Impersonate()

                If impersonationContext IsNot Nothing Then
                    CloseHandle(tokenHandle)
                    Return True
                End If
            End If

            Return False
        End Function

        Public Sub UndoImpersonation()
            If impersonationContext IsNot Nothing Then
                impersonationContext.Undo()
            End If
        End Sub
    End Class

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        Timer1.Start()
        txtIPAddress.Text = My.Settings.IPAddress
        TextBox5.Text = My.Settings.FilePath
        TextBox6.Text = My.Settings.Username
        TextBox7.Text = My.Settings.Password
        TabControl1.Location = New Point(139, 71)
        TabControl1.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        MaximizeBox = False
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
        'Settings Button
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

    Private Sub RichTextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles RichTextBox1.KeyPress
        Dim words As String() = RichTextBox1.Text.Split(" "c)

        ' Check if the number of words exceeds 5
        If words.Length >= 5 AndAlso e.KeyChar <> ControlChars.Back AndAlso Not Char.IsControl(e.KeyChar) AndAlso Char.IsWhiteSpace(e.KeyChar) Then
            e.Handled = True ' Prevent the additional word from being entered
            MessageBox.Show("Word limit exceeded. Please limit your input to 5 words.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
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
        Dim networkPath As String = $"\\{My.Settings.IPAddress}\{My.Settings.FilePath}"
        Dim month As String = Date.Now.ToString("MMMM").ToUpper()
        Dim year As String = Date.Now.Year.ToString()
        Dim txtFileName As String = $"{month}{year}REPORT.txt"
        Dim txtFilePath As String = Path.Combine(networkPath, txtFileName)

        Dim impersonator As New Impersonation()
        Dim isImpersonated As Boolean = False

        Try
            If Not String.IsNullOrEmpty(My.Settings.Username) AndAlso Not String.IsNullOrEmpty(My.Settings.Password) Then
                isImpersonated = impersonator.ImpersonateValidUser(My.Settings.Username, "", My.Settings.Password)
            End If

            If Not Directory.Exists(networkPath) Then
                Directory.CreateDirectory(networkPath)
            End If

            Dim reportData As String = $"Date: {DateTime.Now:yyyy-MM-dd HH:mm tt}" & Environment.NewLine &
                                   $"Ticket Number: {TextBox1.Text}" & Environment.NewLine &
                                   $"Name: {TextBox2.Text}" & Environment.NewLine &
                                   $"Department: {ComboBox1.Text}" & Environment.NewLine &
                                   $"Description: {RichTextBox1.Text}" & Environment.NewLine &
                                   $"Level: {ComboBox2.Text}" & Environment.NewLine &
                                   "-----------------------------------------" & Environment.NewLine

            If File.Exists(txtFilePath) Then
                Dim fileContent As String = File.ReadAllText(txtFilePath)
                If Not fileContent.Contains($"Ticket Number: {TextBox1.Text}") Then
                    Using writer As New StreamWriter(txtFilePath, True, Encoding.UTF8)
                        writer.WriteLine(reportData)
                    End Using
                Else
                    MessageBox.Show("Error: Duplicate ticket number detected. The record was not saved.", "Duplicate Ticket Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            Else
                File.WriteAllText(txtFilePath, reportData)
            End If

            MessageBox.Show("Report data saved to TXT file.")
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error Saving TXT File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If isImpersonated Then
                impersonator.UndoImpersonation()
            End If
        End Try
    End Sub

    Private Sub SaveReportToCSV()
        Dim networkPath As String = $"\\{My.Settings.IPAddress}\{My.Settings.FilePath}"
        Dim year As String = Date.Now.Year.ToString()
        Dim csvFileName As String = $"{year}REPORT.csv"
        Dim csvFilePath As String = Path.Combine(networkPath, csvFileName)

        Dim impersonator As New Impersonation()
        Dim isImpersonated As Boolean = False

        Try
            If Not String.IsNullOrEmpty(My.Settings.Username) AndAlso Not String.IsNullOrEmpty(My.Settings.Password) Then
                isImpersonated = impersonator.ImpersonateValidUser(My.Settings.Username, "", My.Settings.Password)
            End If

            If Not Directory.Exists(networkPath) Then
                Directory.CreateDirectory(networkPath)
            End If

            Dim reportData As String = $"{DateTime.Now:yyyy-MM-dd hh:mm tt},{TextBox1.Text},{TextBox2.Text},{ComboBox1.Text},{RichTextBox1.Text},{ComboBox2.Text}"
            Dim newTicketNumber As String = TextBox1.Text

            Dim existingTickets As New List(Of String)
            If File.Exists(csvFilePath) Then
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

            If existingTickets.Contains(newTicketNumber) Then
                MessageBox.Show("Error: Duplicate ticket number detected. The record was not saved.", "Duplicate Ticket Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If Not File.Exists(csvFilePath) Then
                    Dim header As String = "Date,Ticket Number,Name,Department,Description,Level"
                    File.WriteAllText(csvFilePath, header + Environment.NewLine)
                End If

                File.AppendAllText(csvFilePath, reportData + Environment.NewLine)
                MessageBox.Show("Report data saved to CSV file.")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error Saving CSV File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If isImpersonated Then
                impersonator.UndoImpersonation()
            End If
        End Try
    End Sub

    Private Function GenerateTicketNumber(currentDate As Date, departmentInitial As String) As String
        Dim ticketNumberSuffix As Integer = 1
        Dim ticketNumber As String = currentDate.ToString("ddMMyyyy") & departmentInitial & ticketNumberSuffix

        Dim networkPath As String = $"\\{My.Settings.IPAddress}\{My.Settings.FilePath}"
        Dim year As String = currentDate.Year.ToString()
        Dim csvFileName As String = $"{year}REPORT.csv"
        Dim csvFilePath As String = Path.Combine(networkPath, csvFileName)

        Dim existingTickets As New List(Of String)()

        Try
            Dim impersonation As New Impersonation()
            Dim impersonationSuccessful As Boolean = True

            If Not String.IsNullOrEmpty(My.Settings.Username) AndAlso Not String.IsNullOrEmpty(My.Settings.Password) Then
                impersonationSuccessful = impersonation.ImpersonateValidUser(My.Settings.Username, "", My.Settings.Password)
            End If

            If impersonationSuccessful Then
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
            Else
                MessageBox.Show("Error: Unable to impersonate user.", "Impersonation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            impersonation.UndoImpersonation()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error Generating Ticket Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

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
        Dim networkPath As String = $"\\{My.Settings.IPAddress}\{My.Settings.FilePath}"
        Dim year As String = Date.Now.Year.ToString()
        Dim csvFileName As String = $"{year}REPORT.csv"
        Dim csvFilePath As String = Path.Combine(networkPath, csvFileName)

        Dim impersonator As New Impersonation()
        Dim isImpersonated As Boolean = False

        Try
            If Not String.IsNullOrEmpty(My.Settings.Username) AndAlso Not String.IsNullOrEmpty(My.Settings.Password) Then
                isImpersonated = impersonator.ImpersonateValidUser(My.Settings.Username, "", My.Settings.Password)
            End If

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
                    If Not Directory.Exists(networkPath) Then
                        Directory.CreateDirectory(networkPath)
                    End If

                    Dim pdfFilePath As String = Path.Combine(networkPath, $"{ticketNumber}.pdf")

                    GenerateAndSavePDF(networkPath, ticketNumber, description, departmentSection, reportDate)
                End If
            Else
                MessageBox.Show("Ticket number not found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If isImpersonated Then
                impersonator.UndoImpersonation()
            End If
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
        Dim networkPath As String = $"\\{My.Settings.IPAddress}\{My.Settings.FilePath}"
        Dim pdfFilePath As String = Path.Combine(networkPath, $"{ticketNumber}.pdf")

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
        gfx.DrawString("No: ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(400, 100, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString(ticketNumber, font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(430, 100, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        Dim displayDate As String = ""
        If DateTime.TryParse(reportDate, Nothing) Then
            displayDate = DateTime.Parse(reportDate).ToString("yyyy-MM-dd")
        Else
            displayDate = "Error: Invalid date format"
        End If

        ' Display the date in the PDF
        gfx.DrawString("Date: ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲ ̲", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(400, 170, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString(displayDate, font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(440, 170, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)

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

        ' Draw the trouble description with text wrapping
        Dim italicFont1 As New PdfSharp.Drawing.XFont("Arial", 12, PdfSharp.Drawing.XFontStyle.Italic)
        Dim troubleDescriptionY As Double = troubleBoxY + 15
        Dim maxWidth As Double = troubleBoxWidth - 20
        Dim lineHeight As Double = italicFont1.GetHeight()

        Dim words As String() = description.Split(" "c)
        Dim line As String = ""
        Dim spaceWidth As Double = gfx.MeasureString(" ", font).Width ' Use italic font for measuring space width

        For Each word As String In words
            Dim testLine As String = If(line = "", word, line & " " & word)
            Dim testWidth As Double = gfx.MeasureString(testLine, italicFont1).Width ' Use italic font for measuring width

            If testWidth < maxWidth Then
                line = testLine
            Else
                troubleDescriptionY += lineHeight
                gfx.DrawString(line, italicFont1, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(troubleBoxX + 10, troubleDescriptionY, maxWidth, lineHeight), PdfSharp.Drawing.XStringFormats.TopLeft)
                line = word
            End If
        Next

        ' Draw the remaining text after the loop completes
        If line <> "" Then
            troubleDescriptionY += lineHeight
            gfx.DrawString(line, italicFont1, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(troubleBoxX + 10, troubleDescriptionY, maxWidth, lineHeight), PdfSharp.Drawing.XStringFormats.TopLeft)
        End If

        ' Draw the underline for the description
        Dim underlineHeight As Double = 0 ' Adjust the height of the underline as needed
        Dim underlineY As Double = troubleDescriptionY + lineHeight - underlineHeight
        gfx.DrawRectangle(PdfSharp.Drawing.XPens.Black, troubleBoxX + 10, underlineY, troubleBoxWidth - 20, underlineHeight)

        ' Draw the additional lines
        Dim additionalLineY1 As Double = underlineY + underlineHeight + 20 ' Adjust the spacing between lines as needed
        gfx.DrawRectangle(PdfSharp.Drawing.XPens.Black, troubleBoxX + 10, additionalLineY1, troubleBoxWidth - 20, underlineHeight)

        Dim additionalLineY2 As Double = additionalLineY1 + underlineHeight + 20 ' Adjust the spacing between lines as needed
        gfx.DrawRectangle(PdfSharp.Drawing.XPens.Black, troubleBoxX + 10, additionalLineY2, troubleBoxWidth - 20, underlineHeight)

        Dim additionalLineY3 As Double = additionalLineY2 + underlineHeight + 20 ' Adjust the spacing between lines as needed
        gfx.DrawRectangle(PdfSharp.Drawing.XPens.Black, troubleBoxX + 10, additionalLineY3, troubleBoxWidth - 20, underlineHeight)

        Dim additionalLineY4 As Double = additionalLineY3 + underlineHeight + 20 ' Adjust the spacing between lines as needed
        gfx.DrawRectangle(PdfSharp.Drawing.XPens.Black, troubleBoxX + 10, additionalLineY4, troubleBoxWidth - 20, underlineHeight)


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
        actionDescriptionY += font.GetHeight() + 6
        gfx.DrawString("______________________________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(actionBoxX + 10, actionDescriptionY, actionBoxWidth - 20, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        actionDescriptionY += font.GetHeight() + 6
        gfx.DrawString("______________________________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(actionBoxX + 10, actionDescriptionY, actionBoxWidth - 20, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        actionDescriptionY += font.GetHeight() + 6
        gfx.DrawString("______________________________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(actionBoxX + 10, actionDescriptionY, actionBoxWidth - 20, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        actionDescriptionY += font.GetHeight() + 6
        gfx.DrawString("______________________________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(actionBoxX + 10, actionDescriptionY, actionBoxWidth - 20, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        actionDescriptionY += font.GetHeight() + 6

        ' Draw reported by section
        gfx.DrawString("Reported by:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(40, 330, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("__________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(110, 330, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        Dim reportDateTime As DateTime
        Dim displayTime As String = If(DateTime.TryParse(reportDate, reportDateTime), reportDateTime.ToString("hh:mm tt"), "Error: Invalid date format")
        gfx.DrawString("Time:________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(180, 330, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString(displayTime, font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(210, 330, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Performed by:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(265, 330, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("______________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(340, 330, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Received by:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(40, 355, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("__________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(110, 355, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Approved by:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(40, 375, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("__________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(110, 375, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Accepted by:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(265, 355, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("___________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(335, 355, 200, 0), PdfSharp.Drawing.XStringFormats.TopLeft)

        ' Draw performed by section
        gfx.DrawString("Time Started:________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(430, 330, 50, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString(displayTime, font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(505, 330, 150, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("Time Completed:", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(410, 355, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)
        gfx.DrawString("________", font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XRect(500, 355, 100, 0), PdfSharp.Drawing.XStringFormats.TopLeft)

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
        Process.Start(networkPath)
    End Sub





#End Region

#Region "Tab 3 Code"

    Private Sub txtIPAddress_TextChanged(sender As Object, e As EventArgs) Handles txtIPAddress.TextChanged

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ' Validate the IP address format if needed
        If ValidateIPAddress(txtIPAddress.Text) Then
            ' Save the IP address, file path, username, and password to the settings
            My.Settings.IPAddress = txtIPAddress.Text
            My.Settings.FilePath = TextBox5.Text
            My.Settings.Username = TextBox6.Text
            My.Settings.Password = TextBox7.Text
            My.Settings.Save()
            MessageBox.Show("Settings saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Invalid IP Address format.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Function ValidateIPAddress(ip As String) As Boolean
        ' Simple IP address validation
        Dim ipPattern As String = "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$"
        Return Regex.IsMatch(ip, ipPattern)
    End Function

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

        My.Settings.FilePath = TextBox5.Text
    End Sub


    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        My.Settings.Username = TextBox6.Text
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        Dim password As String = TextBox7.Text
        Dim maskedPassword As String = New String("*"c, password.Length)
        TextBox7.Text = maskedPassword
        TextBox7.Select(password.Length, 0)
    End Sub

    Private Sub TabPage1_Click_1(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

#End Region

End Class
