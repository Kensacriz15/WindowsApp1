Imports MySql.Data.MySqlClient
Imports System.Net.NetworkInformation
Imports System.Collections.Generic
Imports System.Globalization
Module mod_main
    Public openForms As New Dictionary(Of String, Form)()

#Region "CONNECTION STRING"

    Public Sub conString()
        getSet()
        con.Close()
        con.Dispose()

        con.ConnectionString = "SERVER=" & server & ";DATABASE=" & database & ";USERID=" & userid & ";PASSWORD=" & pword & ";PORT=" & port & ";"
    End Sub

    Public Sub conString2()
        getSet()
        con.Close()
        con.ConnectionString = "SERVER=" & server1 & ";DATABASE=" & database1 & ";USERID=" & userid1 & ";PASSWORD=" & pword1 & ";PORT=" & port1 & ";"
    End Sub

    Public Sub conString_res()
        getSet()
        scon.Close()
        scon.Dispose()
        scon.ConnectionString = "SERVER=" & server & ";DATABASE=" & database & ";USERID=" & userid & ";PASSWORD=" & pword & ";PORT=" & port & ";"
    End Sub

#End Region

#Region "SETTINGS"

    Sub getSet()

        Dim AppName As String = Application.ProductName
        Try
            server = GetSetting(AppName, "SystemDB", "DB_IP")
            database = GetSetting(AppName, "SystemDB", "DB_Name")
            userid = GetSetting(AppName, "SystemDB", "DB_User")
            pword = GetSetting(AppName, "SystemDB", "DB_Password")
            port = GetSetting(AppName, "SystemDB", "DB_Port")

            server1 = GetSetting(AppName, "ItemsDB", "DB_IP1")
            database1 = GetSetting(AppName, "ItemsDB", "DB_Name1")
            userid1 = GetSetting(AppName, "ItemsDB", "DB_User1")
            pword1 = GetSetting(AppName, "ItemsDB", "DB_Password1")
            port1 = GetSetting(AppName, "ItemsDB", "DB_Port1")

            If server = "" Then
                server = "192.168.1.231"
                database = "mayer_salesdashboard"
                userid = "root"
                pword = "mayer_admin"
                port = "3306"

                server1 = "192.168.1.231"
                database1 = "mayer_salesdashboard"
                userid1 = "root"
                pword1 = "mayer_admin"
                port1 = "3306"

                saveSet()
            End If
        Catch ex As Exception
            MsgBox("Database con was Not established, you can Set/save ", MsgBoxStyle.Information)
        End Try

    End Sub

    Sub saveSet()

        Dim AppName As String = Application.ProductName

        SaveSetting(AppName, "SystemDB", "DB_IP", server)
        SaveSetting(AppName, "SystemDB", "DB_Name", database)
        SaveSetting(AppName, "SystemDB", "DB_User", userid)
        SaveSetting(AppName, "SystemDB", "DB_Password", pword)
        SaveSetting(AppName, "SystemDB", "DB_Port", port)

        SaveSetting(AppName, "ItemsDB", "DB_IP1", server1)
        SaveSetting(AppName, "ItemsDB", "DB_Name1", database1)
        SaveSetting(AppName, "ItemsDB", "DB_User1", userid1)
        SaveSetting(AppName, "ItemsDB", "DB_Password1", pword1)
        SaveSetting(AppName, "ItemsDB", "DB_Port1", port1)

        MsgBox("Database connection settings are saved. The app will restart", MsgBoxStyle.Information)
        Application.Restart()

    End Sub
    Sub saveSet1()

        Dim AppName As String = Application.ProductName

        SaveSetting(AppName, "SystemDB", "DB_IP", server)
        SaveSetting(AppName, "SystemDB", "DB_Name", database)
        SaveSetting(AppName, "SystemDB", "DB_User", userid)
        SaveSetting(AppName, "SystemDB", "DB_Password", pword)
        SaveSetting(AppName, "SystemDB", "DB_Port", port)

        MsgBox("Database connection settings are saved. The app will restart", MsgBoxStyle.Information)
        Application.Restart()

    End Sub
#End Region

#Region "VARIABLE"
    Public tmeStart As Integer
    Public addFlg As Boolean
    Public cmd As New MySqlCommand
    Public con As New MySqlConnection
    Public scon As New MySqlConnection

    '==========================
    'DASHBOARD LOGIN
    '==========================
    Public uname As String
    Public macaddress As String
    Public pcname As String = Environment.MachineName


    Public database As String = "a"
    Public server As String = "a"
    Public userid As String = "a"
    Public pword As String = "a"
    Public port As String = "a"

    Public database1 As String = "a"
    Public server1 As String = "a"
    Public userid1 As String = "a"
    Public pword1 As String = "a"
    Public port1 As String = "a"

    Public ds As DataSet
    Public ds2 As DataSet
    Public ds3 As DataSet
    Public ds4 As DataSet

    Public da As MySqlDataAdapter
    Public da2 As MySqlDataAdapter
    Public da3 As MySqlDataAdapter
    Public da4 As MySqlDataAdapter

    Public dt As DataTable
    Public dt2 As DataTable
    Public dt3 As DataTable
    Public dt4 As DataTable
    Public dtQ As DataTable

    Public strSQL As String = ""
    Public strSQL2nd As String = ""
    Public conStr As String = ""

    Public m_SortingColumn As ColumnHeader

    Public LV As Boolean = False

    Public ACC_ID As String = ""
    Public ACC_NAME As String = ""

    '==========added for MayerAdmin==========
    '
    'Public itmcode As String = ""
    'Public itmcode1 As String = ""
    'Public id As String = ""
    'Public id1 As String = ""

    '=========added for MayerPayroll=========
    Public user As String = ""
    Public vstring As String = ""
    Public vstring1 As String = ""
    Public vstring2 As String = ""
    Public vstring3 As String = ""
    Public vstring4 As String = ""
    Public LV1 As Boolean = True
    Public empno As String
    Public notif_access As Integer = 0
    Public con_success As Boolean = True 'For invalid connections logged as administrator; Used for Restriction of module

#End Region


#Region "QUERY"

    Public strCollect As String = ""
    Public Sub sqlCollect(ByVal str As String)
        If str.Length > 0 Then
            If str.Substring(str.Length - 1, 1) <> ";" Then
                str = str & ";"
            End If
        End If
        strCollect = strCollect & str
    End Sub

    Function ExecuteQueryaa(ByVal strSQL As String, ByVal dsTable As String) As DataTable

        'Try
        If con.State = ConnectionState.Closed Then
            conString()
            con.Open()
        End If

        ds = New DataSet
        da = New MySqlDataAdapter(strSQL, con)
        da.Fill(ds, dsTable)

        'Catch ex As MySql.Data.MySqlClient.MySqlException
        'MsgBox(ex.Message)
        '    constr()
        'End Try

        Return ds.Tables(dsTable)

    End Function

    Public Sub _at(uname As String, pcname As String, macadd As String, frmname As String, ctrlname As String, rmrks As String)
        SaveRecord("INSERT INTO tbl_audittrail(nuser,npcname,macadd,dttrans,frm,frmcntrl,remarks) VALUES('" & uname & "', '" & pcname & "', '" & macadd & "',Now(), '" & frmname & "', '" & ctrlname & "', '" & rmrks & "');")
    End Sub
    Public Sub SaveRecord(ByVal sqlStr As String)

        Try
            If con.State = ConnectionState.Closed Then
                conString()
                con.Open()
            End If
            'MsgBox(sqlStr)
            cmd = New MySqlCommand
            cmd.CommandTimeout = 600
            cmd.Connection = con
            cmd.CommandText = sqlStr
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Function ExecuteQuery(ByVal strSQL As String) As DataTable

        Dim dsTable As String = "dsTable"

        If con.State = ConnectionState.Closed Then
            conString()
            con.Open()
        End If

        ds = New DataSet
        da = New MySqlDataAdapter(strSQL, con)
        da.Fill(ds, dsTable)

        Return ds.Tables(dsTable)
    End Function
    Function ExecuteQueryPrint(ByVal strSQL As String) As DataTable

        Dim dsTable As String = "dtDailySales"

        If con.State = ConnectionState.Closed Then
            conString()
            con.Open()
        End If

        ds = New DataSet
        da = New MySqlDataAdapter(strSQL, con)
        da.Fill(ds, dsTable)

        Return ds.Tables(dsTable)
    End Function
    Function getMacAddress()
        Dim nics() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces
        Return nics(0).GetPhysicalAddress.ToString
    End Function
    Function ExecuteQueries(ByVal strSQL As String) As DataTable

        Dim dsTable As String = "dsTable"

        If scon.State = ConnectionState.Closed Then
            conString_res()
            scon.Open()
        End If

        ds = New DataSet
        da = New MySqlDataAdapter(strSQL, scon)
        da.Fill(ds, dsTable)
        Return ds.Tables(dsTable)

    End Function

    Function ExecuteQuery2(ByVal strSQL As String) As DataTable

        Dim dsTable As String = "dsTable2"
        If con.State = ConnectionState.Closed Then
            conString2()
            con.Open()
        End If

        ds2 = New DataSet
        da2 = New MySqlDataAdapter(strSQL, con)
        da2.Fill(ds2, dsTable)
        Return ds2.Tables(dsTable)

    End Function

    Function clrstr(ByVal st As String)
        st = Replace(st, "'", "''")
        st = Replace(st, "\", "\\")
        Return st
    End Function

    Public Sub cboQuery(ByVal col As String, ByVal tbl As String, ByVal cbo As ComboBox)
        Try
            strSQL = "SELECT " & col & " FROM " & tbl & " GROUP BY " & col & ";"
            dt = ExecuteQuery(strSQL)
            cbo.Items.Clear()
            If dt.Rows.Count > 0 Then
                For q As Integer = 0 To dt.Rows.Count - 1
                    cbo.Items.Add(dt.Rows(q)(0).ToString)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "GET_TRANSACTIONNO"
    Function GET_TR_NO(ByVal col As String, ByVal TrKey As String) As String
        Dim tr_counter As Integer = 0
        Try
            strSQL = "SELECT " & col & " FROM `ps_generator` ;"
            dt = ExecuteQuery(strSQL)
            If dt.Rows.Count > 0 Then
                tr_counter = Val(dt.Rows(0)(0).ToString)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        tr_counter = tr_counter + 1
        Return TrKey + "-" + Format(tr_counter, "000000")
    End Function

    Sub SET_TR_NO(ByVal col As String)
        SaveRecord("UPDATE ps_generator SET " & col & "=" & col & "+1 ;")
    End Sub
#End Region

#Region "SORT COLUMN"
    Public Sub SortColumn(ByVal lv As ListView, ByVal e As System.Windows.Forms.ColumnClickEventArgs)
        ' Get the new sorting column.
        Dim new_sorting_column As ColumnHeader =
            lv.Columns(e.Column)
        ' Figure out the new sorting order.
        Dim sort_order As System.Windows.Forms.SortOrder
        If m_SortingColumn Is Nothing Then
            ' New column. Sort ascending.
            sort_order = SortOrder.Ascending
        Else
            ' See if this is the same column.
            If new_sorting_column.Equals(m_SortingColumn) Then
                ' Same column. Switch the sort order.
                If m_SortingColumn.Text.StartsWith("> ") Then
                    sort_order = SortOrder.Descending
                Else
                    sort_order = SortOrder.Ascending
                End If
            Else
                ' New column. Sort ascending.
                sort_order = SortOrder.Ascending
            End If

            ' Remove the old sort indicator.
            m_SortingColumn.Text =
                m_SortingColumn.Text.Substring(2)
        End If
        ' Display the new sort order.
        m_SortingColumn = new_sorting_column
        If sort_order = SortOrder.Ascending Then
            m_SortingColumn.Text = "> " & m_SortingColumn.Text
        Else
            m_SortingColumn.Text = "< " & m_SortingColumn.Text
        End If
        ' Create a comparer.
        lv.ListViewItemSorter = New _
            ListViewComparer(e.Column, sort_order)
        ' Sort.
        lv.Sort()
    End Sub
#End Region

#Region "NUMBER"

    Public Sub numkeys(ByVal txtbox As TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Select Case Asc(e.KeyChar)
            Case 8, 46, 48 To 57
                If Not txtbox.Text.IndexOf(".") = -1 AndAlso e.KeyChar = "." Then e.Handled = True
            Case Else
                e.Handled = True
        End Select
    End Sub

    Public Sub numkeys1(ByVal txtbox As ToolStripTextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Select Case Asc(e.KeyChar)
            Case 8, 46, 48 To 57
                If Not txtbox.Text.IndexOf(".") = -1 AndAlso e.KeyChar = "." Then e.Handled = True
            Case Else
                e.Handled = True
        End Select
    End Sub

    Public Sub keynum(ByVal txt As ToolStripTextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim keyascii As Integer = Asc(e.KeyChar)
        If keyascii = 8 Then
            keyascii = 8
            Exit Sub
        End If
        If keyascii < 48 Or keyascii > 57 Then e.Handled = True
    End Sub

    '======================Added for MayerPayroll==============
    'Whole number
    Public Sub keynum1(ByVal txt As TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        Dim keyascii As Integer = Asc(e.KeyChar)
        If keyascii = 8 Then
            keyascii = 8
            Exit Sub
        End If
        If keyascii < 48 Or keyascii > 57 Then e.Handled = True
    End Sub

#End Region

#Region "Get Version"
    Function GetVersion() As String
        With System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location)
            Return .FileMajorPart & "." & .FileMinorPart & "." & .FileBuildPart & "." & .FilePrivatePart
        End With
    End Function
#End Region

    Public weekDateRanges_from As New List(Of String)
    Public weekDateRanges_to As New List(Of String)
    Function GetWeekDateRanges(year As Integer, month As Integer) As List(Of String)
        Dim calendar As Calendar = CultureInfo.CurrentCulture.Calendar
        Dim dateRanges As New List(Of String)

        ' Get the first day of the month
        Dim currentDate As New DateTime(year, month, 1)

        ' Loop through the weeks in the month
        While currentDate.Month = month
            Dim endOfWeek As DateTime = currentDate.AddDays(7 - CInt(currentDate.DayOfWeek))

            ' Adjust for the end of the month
            If endOfWeek.Month > month Then
                endOfWeek = New DateTime(year, month, DateTime.DaysInMonth(year, month))
            End If

            ' Add the date range to the list
            dateRanges.Add($"{currentDate: MM/dd/yyyy} And {endOfWeek: MM/dd/yyyy}")
            weekDateRanges_from.Add($"{currentDate: MM/dd/yyyy}")
            weekDateRanges_to.Add($"{endOfWeek: MM/dd/yyyy}")

            ' Move to the next week
            currentDate = endOfWeek.AddDays(1)
        End While


        Return dateRanges
    End Function
End Module
Class ListViewComparer
    Implements IComparer

    Private m_ColumnNumber As Integer
    Private m_SortOrder As SortOrder

    Public Sub New(ByVal column_number As Integer, ByVal _
        sort_order As SortOrder)
        m_ColumnNumber = column_number
        m_SortOrder = sort_order
    End Sub

    ' Compare the items in the appropriate column
    ' for objects x and y.
    Public Function Compare(ByVal x As Object, ByVal y As _
        Object) As Integer Implements _
        System.Collections.IComparer.Compare
        Dim item_x As ListViewItem = DirectCast(x,
            ListViewItem)
        Dim item_y As ListViewItem = DirectCast(y,
            ListViewItem)

        ' Get the sub-item values.
        Dim string_x As String
        If item_x.SubItems.Count <= m_ColumnNumber Then
            string_x = ""
        Else
            string_x = item_x.SubItems(m_ColumnNumber).Text
        End If

        Dim string_y As String
        If item_y.SubItems.Count <= m_ColumnNumber Then
            string_y = ""
        Else
            string_y = item_y.SubItems(m_ColumnNumber).Text
        End If

        ' Compare them.
        If m_SortOrder = SortOrder.Ascending Then
            Return String.Compare(string_x, string_y)
        Else
            Return String.Compare(string_y, string_x)
        End If
    End Function

End Class