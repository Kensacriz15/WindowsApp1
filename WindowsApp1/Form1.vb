Imports System.IO


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

    Public Function GenerateUnique5DigitNumber() As Long
        Dim random As New Random()
        Dim newNumber As Long

        Do
            newNumber = random.Next(10001, 100000)
        Loop While UsedNumbers.Contains(newNumber)

        UsedNumbers.Add(newNumber)
        Return newNumber
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim randomNumber As Long = GenerateUnique5DigitNumber()
        TextBox1.Text = randomNumber.ToString()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'Generate Ticket Number
        TextBox1.ReadOnly = True
    End Sub

    Private originalText As String

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

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

        If defaultSelection Then
            defaultSelection = False


            ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
            ComboBox1.SelectedIndex = 0 '
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        'Description
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim filePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\test.txt"
        Dim data As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "|" & TextBox1.Text & "|" & TextBox2.Text & "|" & ComboBox1.Text & "|" & ComboBox2.Text & "|" & RichTextBox1.Text

        If File.Exists(filePath) Then
            File.AppendAllText(filePath, Environment.NewLine & data)
        Else
            File.WriteAllText(filePath, data)
        End If

        Dim values() As String = File.ReadAllLines(filePath)
        Dim lastEntry As String = values(values.Length - 1)
        Dim lastEntryValues() As String = lastEntry.Split("|"c)
        TextBox1.Text = lastEntryValues(1)
        TextBox2.Text = lastEntryValues(2)
        ComboBox1.Text = lastEntryValues(3)
        ComboBox2.Text = lastEntryValues(4)
        RichTextBox1.Text = lastEntryValues(5)
        MessageBox.Show("Report Ticket Submitted")

    End Sub

    Private Sub RichTextBox2_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox2.TextChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim filePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\test.txt"
        Dim fileContent As String = File.ReadAllText(filePath)
        RichTextBox2.Text = fileContent

        Dim timer As New Timer()
        timer.Interval = 2000 ' 
        AddHandler timer.Tick, AddressOf RefreshFileContent
        timer.Start()
    End Sub

    Private Sub RefreshFileContent(sender As Object, e As EventArgs)
        Dim filePath As String = "C:\Users\MIS - Rafael\Desktop\PLDT\test.txt"
        Dim fileContent As String = File.ReadAllText(filePath)
        RichTextBox2.Text = fileContent
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Console.WriteLine(DateAndTime.Now)
        Console.ReadLine()
    End Sub
End Class