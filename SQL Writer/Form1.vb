Option Explicit On
Option Strict On

Imports System.IO
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel


Public Class frmMain

    Dim myFile As New FilePath
    Dim myText As New TextVB
    Dim help As New frmHelp
    Dim mSQL As New frmSQL
    Dim myImport As New frmImport
    Dim myfrmTable As New frmTableFields
    Dim myScript As String = ""


    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        myFile.File = ""
        myFile.Script = ""
        cbxSelect.Checked = True
        ClearItems()
        AutoScroll = True
        DisplayVersion()
        lblAuthor.Text = "Author: Lonnie Coates, Jr."
        lblAuthor.Refresh()

        KeyPreview = True
    End Sub
    Private Sub DisplayVersion()
        Dim strVersionType As String

        strVersionType = " Release"
        Dim strVersionNo As String

        strVersionNo = Application.ProductVersion.ToString
        lblVersion.Text = "V. " & strVersionNo & strVersionType

        lblVersion.Refresh()
    End Sub

    Private Function SaveFileWindow() As String
        Dim MyAppDir As String = Environment.CurrentDirectory

        Dim strFile As String
        Dim intAnswer As Integer

        intAnswer = 0
        strFile = ""

        Dim saveFileDialog1 As New SaveFileDialog()

        saveFileDialog1.InitialDirectory = MyAppDir

        saveFileDialog1.Filter = "Access 2007 file (*.Accdb)|*.Accdb |Access 2003 file (*.mdb)|*.mdb"

        saveFileDialog1.FilterIndex = 1
        saveFileDialog1.RestoreDirectory = True

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then

            strFile = saveFileDialog1.FileName
        End If

        Return strFile
    End Function
    Private Function Connection() As String ''Not Used
        Dim MyAppDir As String = Environment.CurrentDirectory

        Dim dbFile As String = "\" & "ArtDB.Accdb"

        Return MyAppDir & dbFile
    End Function

    Private Sub DataGridView1_MouseWheel(sender As Object, e As System.Windows.Forms.MouseEventArgs)
        'Moves scrollbar with mousewheel
        Dim intMove As Integer = System.Convert.ToInt32(DataGridView1.FirstDisplayedScrollingRowIndex - e.Delta / 120)

        If (intMove >= 0) And (intMove <= DataGridView1.RowCount) Then
            DataGridView1.FirstDisplayedScrollingRowIndex = intMove
        End If

    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        CloseDatabase()

        Application.Exit()
    End Sub

    Private Sub btnCreateDB_Click(sender As Object, e As EventArgs) Handles btnCreateDB.Click
        Dim strFile As String = ""
        strFile = SaveFileWindow()

        If strFile.Length = 0 Then
        Else
            SaveDB(strFile)
        End If
    End Sub
    Private Sub SaveDB(ByVal strFile As String)
        ''The older Microsoft "Jet" drivers are only available to 32-bit applications. 
        ''If you want to use Jet you'll need to go into the "Build" tab of the Properties...
        ''...for your Visual Studio project and set the "Target platform" to "x86". 
        ''That will force your application to run as 32-bit. See the link below:
        ''Link to article: stackoverflow.com/a/28049916
        Dim bAns As Boolean
        Dim cat As New ADOX.Catalog()
        Try

            Dim sCreateString As String
            sCreateString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile & ";Jet OLEDB:Engine Type=5"

            cat.Create(Convert.ToString(sCreateString))

            bAns = True
            cat = Nothing
        Catch Excep As Exception
            bAns = False
            MsgBox(Excep.Message, MsgBoxStyle.Information, "Error")
        Finally
            cat = Nothing
        End Try
    End Sub

    Private Sub OpenScript()

        Dim strFile As String = ""
        Dim strFilePath As String = ""
        Dim strFileName As String = ""
        Dim strFilters As String = ""
        Dim inAnswer As Integer = 0
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim strDefaultPath As String = ""
        Dim strFileText As String = ""

        If myFile.Script.Length > 4 Then
            strFileName = myFile.Script.Remove(0, InStrRev(myFile.Script, "\"))
            strFilePath = myFile.Script.Replace(strFileName, "")
        End If
        ''
        If strFilePath.Length > 0 Then
            strDefaultPath = strFilePath
        Else
            strDefaultPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End If
        OpenFileDialog1.InitialDirectory = strDefaultPath

        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Text File|*.txt"
        OpenFileDialog1.Multiselect = False

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            myFile.Script = OpenFileDialog1.FileName
            Dim streamRead As New StreamReader(OpenFileDialog1.FileName)
            strFileText = streamRead.ReadToEnd
            streamRead.Close()

            rtxtSelect.Text = strFileText
            If strFileText.Equals("") Then
                MsgBox("There's nothing in this file.", MsgBoxStyle.Information)
            End If
        End If
    End Sub

    Private Sub SaveScript()
        Dim intAnswer As Integer = 0
        Dim strFile As String = ""
        Dim strFilePath As String = ""
        Dim strFileName As String = ""
        Dim strFilters As String = ""
        Dim inAnswer As Integer = 0
        Dim SaveFile As New SaveFileDialog
        Dim strDefaultPath As String = ""
        Dim strFileText As String = ""

        If myFile.Script.Length > 4 Then
            strFileName = myFile.Script.Remove(0, InStrRev(myFile.Script, "\"))
            strFilePath = myFile.Script.Replace(strFileName, "")
        End If
        ''
        If strFilePath.Length > 0 Then
            strDefaultPath = strFilePath
        Else
            strDefaultPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End If
        SaveFile.InitialDirectory = strDefaultPath

        SaveFile.FileName = ""
        SaveFile.Filter = "Text File|*.txt"

        If SaveFile.ShowDialog() = DialogResult.OK Then
            myFile.Script = SaveFile.FileName
            Dim streamWrite As New StreamWriter(SaveFile.FileName)
            strFileText = rtxtSelect.Text
            streamWrite.WriteLine(strFileText)
            streamWrite.Close()

            intAnswer = MsgBox("The script for " & myFile.Script.Remove(0, InStrRev(myFile.Script, "\")) & " was successfully saved." _
                   & " Would you like to open the file?", MsgBoxStyle.YesNo, "Save successful.")
            If intAnswer = vbYes Then
                Try
                    Process.Start(SaveFile.FileName.ToString)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error")
                End Try

            End If
        End If
    End Sub

    Private Sub OpenDB()

        Dim strFile As String = ""
        Dim strFilePath As String = ""
        Dim strFileName As String = ""
        Dim strFilters As String = ""
        Dim inAnswer As Integer = 0
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim strDefaultPath As String = ""
        If myFile.File.Length > 4 Then
            strFileName = myFile.File.Remove(0, InStrRev(myFile.File, "\"))
            strFilePath = myFile.File.Replace(strFileName, "")
        End If
        ''
        If strFilePath.Length > 0 Then
            strDefaultPath = strFilePath
        Else
            strDefaultPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End If
        OpenFileDialog1.InitialDirectory = strDefaultPath

        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Access 2007 file (*.Accdb)|*.Accdb|Access 2003 file (*.mdb)|*.mdb|All Access filetypes|*.Accdb; *.mdb"
        OpenFileDialog1.Multiselect = False

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            strFile = OpenFileDialog1.FileName
            ClearItems()
            LoadTables(strFile)
            myFile.File = strFile
            FileTrim(strFile, 106)
            tTip.SetToolTip(lblFilePath, myFile.File)
            lblFilePath.Text = strFile
            lblFilePath.Refresh()
        End If
    End Sub

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        RunButton()
    End Sub

    Private Sub RunButton()
        If File.Exists(myFile.File) Then
            If rtxtSelect.Text.Length > 0 Then
                If cbxSelect.Checked = True Then
                    If rtxtSelect.Text.ToLower.Contains("select") = False Then
                        MsgBox("You didn't write a 'SELECT' statement.")
                    Else
                        Try
                            myScript = rtxtSelect.Text
                            RunScript(myScript, myFile.File)
                        Catch ex As Exception
                            MsgBox(ex.Message, MsgBoxStyle.Information, "Error")
                        End Try
                    End If
                Else

                    Try
                        RunScript(rtxtSelect.Text, myFile.File, 1)
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Information, "Error")
                    End Try

                End If
            Else
                MsgBox("You didn't write a SQL statement.", MsgBoxStyle.Information, "Error")
            End If
        Else
            MsgBox("You didn't select an Access database file.", MsgBoxStyle.Information, "Error")
        End If
    End Sub

    Private Sub RunScript(ByVal strScript As String, ByVal thisFile As String, ByVal Optional intType As Integer = 0)
        Dim con1 As New OleDbConnection
        Dim boolError As Boolean = False
        Dim errorMessage As String = ""
        con1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & thisFile
        Dim sqlNo As Integer = 0
        Dim sql As String
        sql = strScript

        If intType = 0 Then
            Dim adapter As New OleDbDataAdapter(sql, con1)
            ' Gets the records from the table and fills our adapter with those.
            Dim dt As New DataTable("tblProjects")
            adapter.Fill(dt)
            ' Assigns our DataSource on the DataGridView
            DataGridView1.DataSource = dt

            DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Else
            ''Testing SQL statements to run correctly before commiting to even the first statement
            errorMessage = TestScript(strScript, thisFile)
            If errorMessage.Length > 0 Then
                boolError = True
                MsgBox(errorMessage, MsgBoxStyle.Information, "SQL Statement Error")
            Else

                con1.Open()
                Dim SemiColonCount As Integer = 0
                SemiColonCount = strScript.Count(Function(c As Char) c = ";")
                Dim strFullScript As String = ""

                ''Dim SemiColonCount2 As Integer = 0
                strFullScript = strScript.Replace(vbCr, "")
                strFullScript = strFullScript.Replace(vbLf, "")

                If SemiColonCount > 1 Then
                    Try
                        For i As Integer = 0 To SemiColonCount - 1 Step 1
                            sql = strFullScript.Substring(0, InStr(1, strFullScript, ";"))
                            Dim count As Integer = InStr(1, strFullScript, ";")

                            sqlNo = i + 1
                            Dim cmd As New OleDbCommand(sql, con1)
                            cmd.ExecuteNonQuery()
                            strFullScript = strFullScript.Remove(0, count)

                        Next
                    Catch dbException As OleDbException
                        MsgBox(dbException.Message + " Error on statement No. " + sqlNo.ToString(),
                            MsgBoxStyle.Information, "Multi-Statement Error")
                        If con1.State.ToString.ToLower.Equals("open") Then
                            con1.Close()
                        End If

                        boolError = True
                    End Try

                Else
                    Dim cmd As New OleDbCommand(sql, con1)
                    cmd.ExecuteNonQuery()
                End If
            End If
            sql = strScript


            ClearItems()
            lblFilePath.Text = thisFile
            lstTables.Refresh()
            LoadTables(thisFile)

            If boolError = False Then
                MsgBox("SQL statement: " & sql & " was successfully executed.", MsgBoxStyle.Information, "SQL Message")
            End If
        End If

        If con1.State.ToString.ToLower.Equals("open") Then
            con1.Close()
        End If
    End Sub

    ''Tests to make sure that all of the SQL statements run correctly before commiting to even the first statement
    Private Function TestScript(ByVal strScript As String, ByVal thisFile As String) As String
        Dim con1 As New OleDbConnection
        Dim transaction1 As OleDbTransaction = Nothing
        Dim boolError As Boolean = False
        con1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & thisFile
        Dim sqlNo As Integer = 0
        Dim sql As String = ""
        Dim errorMessage As String = ""
        sql = strScript
        con1.Open()
        Dim SemiColonCount As Integer = 0
        SemiColonCount = strScript.Count(Function(c As Char) c = ";")
        Dim strFullScript As String = ""

        strFullScript = strScript.Replace(vbCr, "")
        strFullScript = strFullScript.Replace(vbLf, "")

        If SemiColonCount > 0 Then
            Try
                For i As Integer = 0 To SemiColonCount - 1 Step 1
                    sql = strFullScript.Substring(0, InStr(1, strFullScript, ";"))

                    Dim count As Integer = InStr(1, strFullScript, ";")

                    sqlNo = i + 1

                    transaction1 = con1.BeginTransaction()

                    Dim cmd As New OleDbCommand() ''sql, con1) ''
                    cmd.CommandText = sql
                    cmd.Connection = con1
                    cmd.Transaction = transaction1
                    Dim intRows As Integer = cmd.ExecuteNonQuery()

                    transaction1.Rollback()

                    strFullScript = strFullScript.Remove(0, count)

                Next
            Catch dbException As OleDbException
                Try
                    transaction1.Rollback()
                Catch ex As Exception
                End Try
                errorMessage = dbException.Message + " Error on statement No. " + sqlNo.ToString()

                If con1.State.ToString.ToLower.Equals("open") Then
                    con1.Close()
                End If

                boolError = True
            End Try
        Else
            Try


                Dim cmd As New OleDbCommand(sql & " SET NOEXEC ON;", con1)
                cmd.ExecuteNonQuery()
            Catch dbException As OleDbException
                errorMessage = dbException.Message + " Error on statement No. 1"
                If con1.State.ToString.ToLower.Equals("open") Then
                    con1.Close()
                End If
            End Try
        End If
        If con1.State.ToString.ToLower.Equals("open") Then
            con1.Close()
        End If

        Return errorMessage
    End Function

    Private Sub ClearItems()
        lblFilePath.Text = ""
        tTip.SetToolTip(lblFilePath, "")
        lstTables.Items.Clear()
        lstTables.Refresh()
    End Sub

    Private Sub LoadTables(ByVal thisFile As String)
        Dim userTables As DataTable = Nothing
        Dim con1 As New OleDbConnection

        Dim restrictions() As String = New String(3) {}
        Try
            restrictions(3) = "Table"
            ''restrictions(3) = "Query"
            con1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & thisFile
            con1.Open()
            userTables = con1.GetSchema("Tables", restrictions)
            con1.Close()

            For i As Integer = 0 To userTables.Rows.Count - 1 Step i + 1
                lstTables.Items.Add(userTables.Rows(i)(2).ToString)
            Next
            lstTables.Refresh()
        Catch ex As Exception

            con1.Close()

        End Try
        If con1.State.ToString.ToLower.Equals("open") Then
            con1.Close()
        End If
    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        Try
            'Opens Access DB file
            OpenDB()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error")
        End Try
    End Sub

    Private Sub lstTables_DoubleClick(sender As Object, e As EventArgs) Handles lstTables.DoubleClick
        ''Adds selected table to the SQL textbox
        rtxtSelect.Text = rtxtSelect.Text + lstTables.SelectedItem.ToString
    End Sub

    Private Sub btnSaveScript_Click(sender As Object, e As EventArgs) Handles btnSaveScript.Click
        Try
            ''Saves TXT script file
            SaveScript()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Error")
        End Try
    End Sub

    Private Sub btnLoadScript_Click(sender As Object, e As EventArgs) Handles btnLoadScript.Click
        Try
            ''Loads Saved TXT script file
            OpenScript()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Error")
        End Try
    End Sub

    Private Sub btnSaveReport_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveReport.Click
        Dim strFile As String = ""

        'where the user wants to save the .xls file
        strFile = SaveReportWindow()
        Dim officeType As Type = Type.GetTypeFromProgID("Excel.Application")
        If DataGridView1.ColumnCount > 0 Then

            If IsNothing(officeType) = True Then
                'Doesn't Require MS Office to Work
                SaveXLS(strFile)
            Else
                'Requires MS Office to Work
                SaveToExcel(strFile)
            End If

        Else
            MsgBox("There's no data to save the report with.", MsgBoxStyle.Information, "Report Error")
        End If
    End Sub


    'Requires MS Office to work
    Private Sub SaveToExcel(ByVal strFile As String)
        Dim oExcel As Excel.Application = Nothing
        Dim oBook As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim oRange As Excel.Range = Nothing

        Try

            If strFile.Equals("") Or strFile = Nothing Then

            Else
                'Create New Excel File
                oExcel = New Excel.Application
                'oExcel.enable
                oExcel.Visible = False

                'Create New Workbook
                oBook = oExcel.Workbooks.Add

                oSheet = DirectCast(oBook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet) 'DirectCast necessary when Option Strict is on

                Dim thisScript As String = ""

                thisScript = myScript.Replace(vbLf, "")
                Dim intTotalColumn, intTotalRow As Integer
                Dim intColIndex, intRowIndex As Integer
                Dim strResults As String = "SQL Statement:  " & thisScript
                intTotalColumn = DataGridView1.ColumnCount

                intTotalRow = DataGridView1.RowCount

                DirectCast(oSheet.Cells(intRowIndex + 1, intColIndex + 1), Excel.Range).Value = strResults

                For intRowIndex = 0 To intTotalRow - 1
                    For intColIndex = 0 To intTotalColumn - 1
                        DirectCast(oSheet.Cells(intRowIndex + 3, intColIndex + 1), Excel.Range).Value =
                            DataGridView1.Item(intColIndex, intRowIndex).Value

                        If intRowIndex = 0 Then
                            DirectCast(oSheet.Cells(intRowIndex + 2, intColIndex + 1), Excel.Range).Value =
                                DataGridView1.Columns(intColIndex).Name
                        End If


                    Next

                    If intRowIndex = 0 Then
                        'Formats Title of Table
                        oRange = oSheet.Range(oSheet.Cells(intRowIndex + 1, 1), oSheet.Cells(intRowIndex + 1, intColIndex))
                        oRange.Font.Bold = True
                        oRange.MergeCells = True
                        oRange.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter


                        'Formats Table Fields
                        oRange = oSheet.Range(oSheet.Cells(intRowIndex + 2, 1), oSheet.Cells(intRowIndex + 2, intColIndex))
                        oRange.Font.Bold() = True


                    End If
                Next

                oExcel.DisplayAlerts = False

                'Auto Sizes all columns
                oSheet.Columns.AutoFit()


                oBook.SaveAs(strFile, -4143) ' -4143 = Normal Workbook

                oBook.Close()
                oExcel.Quit()

                Dim intAnswer As Integer = 0
                intAnswer = MessageBox.Show("File was success exported to " & strFile & "." & Chr(13) & Chr(13) & "Do you wish open the file?", "Export Successful",
                                            MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                'If user wants to open the newly created file....
                If intAnswer = vbYes Then

                    '...open file with default program
                    Try
                        Process.Start(strFile)
                    Catch

                    End Try
                End If
            End If
        Catch ex As Exception
            oBook.Close()
            oExcel.Quit()
        End Try

    End Sub


    'Saves search results as an xls excel file without Excel being installed
    Private Sub SaveXLS(ByVal strFile As String)

        Dim myCommand As New OleDbCommand
        Dim strSearch, strSQLCommand, strColums As String
        Dim intColumns As Integer = 0
        Dim i As Integer = 0
        Dim intRows As Integer = 0
        Dim strRows As String = ""



        'Creates the connection
        Dim con1 As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFile & ";Extended Properties=Excel 8.0;")
        strSearch = "SQL Results"


        intColumns = DataGridView1.ColumnCount

        If strFile.Equals("") Or strFile = Nothing Then
        Else
            strColums = ""

            For i = 0 To intColumns - 1 Step 1
                If i > 0 Then
                    strColums = strColums & ", "
                End If
                Dim datatype As String = "char(200)"
                Dim valueType As String
                valueType = DataGridView1.Columns(i).ValueType.ToString

                If valueType.Equals("system.int32") Then
                    datatype = "int"
                ElseIf valueType.ToLower.Equals("system.Double") Then
                    datatype = "Double"
                ElseIf valueType.ToLower.Equals("system.datetime") Then
                    datatype = "datetime"
                End If

                'Gets a collection of column names from data table in the data grid view control
                strColums = strColums & " [" & DataGridView1.Columns(i).Name.ToString & "] " & datatype

            Next

            'Tab name will be the value of strSearch (the user's search results)
            strSQLCommand = "Create Table [" & strSearch & "] (" & strColums & ")"



            Try
                'If user wants to override the current file (if it exist)...
                If File.Exists(strFile) Then
                    '...delete old version of that file
                    File.Delete(strFile)
                End If
                con1.Open()
                myCommand.Connection = con1

                myCommand.CommandText = strSQLCommand
                myCommand.ExecuteNonQuery()

                intRows = DataGridView1.RowCount
                For index As Integer = 0 To intRows - 1 Step 1
                    strRows = ""
                    For i = 0 To intColumns - 1 Step 1
                        If i > 0 Then
                            strRows = strRows & ", "
                        End If
                        'Stores a collection of data to go under the Columns
                        Dim tempValue As String = DataGridView1.Item(i, index).Value.ToString


                        If IsNumeric(tempValue) = True Then
                            strRows = strRows & DataGridView1.Item(i, index).Value.ToString
                        ElseIf IsDate(tempValue) = True Then
                            strRows = strRows & "#" & DataGridView1.Item(i, index).Value.ToString & "#"
                        Else
                            strRows = strRows & "'" & DataGridView1.Item(i, index).Value.ToString & "'"
                        End If


                    Next
                    strSQLCommand = "INSERT INTO [" & strSearch & "] (" & strColums.Replace(" char(200)", "") & ") Values(" & strRows & ")"

                    myCommand.CommandText = strSQLCommand
                    myCommand.ExecuteNonQuery()
                Next

                Dim intAnswer As Integer = 0
                intAnswer = MessageBox.Show("File was success exported to " & strFile & "." & Chr(13) & Chr(13) & "Do you wish open the file?", "Export Successful",
                                            MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                'If user wants to open the newly created file....
                If intAnswer = vbYes Then
                    con1.Close()

                    '...open file with default program
                    Try
                        Process.Start(strFile)
                    Catch

                    End Try
                End If

            Catch exception As System.Exception
                MessageBox.Show(exception.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            End Try
        End If
        con1.Close()
    End Sub

    Private Function SaveReportWindow() As String
        Dim MyAppDir As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        Dim strFile As String
        Dim intAnswer As Integer

        intAnswer = 0

        strFile = ""

        Dim saveFileDialog1 As New SaveFileDialog()

        saveFileDialog1.InitialDirectory = MyAppDir

        saveFileDialog1.Filter = "Excel 2003 / OpenOffice Cal (*.xls)|*.xls"
        saveFileDialog1.FilterIndex = 1
        saveFileDialog1.RestoreDirectory = True

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then

            strFile = saveFileDialog1.FileName
        End If

        Return strFile
    End Function

    Private Sub btnHelp_Click(sender As Object, e As EventArgs) Handles btnHelp.Click
        Dim frmCollection = System.Windows.Forms.Application.OpenForms


        If frmCollection.OfType(Of frmHelp).Any Then
            Me.BringToFront()
            help.BringToFront()
        Else
            help = New frmHelp
            help.Show()
        End If
    End Sub

    Private Sub btnCloseDB_Click(sender As Object, e As EventArgs) Handles btnCloseDB.Click

        CloseDatabase()

    End Sub
    ''Closes Database file
    Private Sub CloseDatabase()
        ''myFile.File = ""
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        ClearItems()
    End Sub

    Private Sub btnTableFields_Click(sender As Object, e As EventArgs) Handles btnTableFields.Click
        Dim frmCollection = System.Windows.Forms.Application.OpenForms
        Dim strTableName As String

        ''Check to see if lstTables is empty
        If IsNothing(lstTables.SelectedItem) = True Then
            strTableName = ""
        Else
            ''Tries to set selected row from lstTables into a variable
            strTableName = lstTables.SelectedItem.ToString
        End If

        If strTableName.Equals("") Then
            MsgBox("You didn't select a table.", MsgBoxStyle.Information)
        Else
            ''Brings the form for Table fields to the front
            If frmCollection.OfType(Of frmTableFields).Any Then
                Me.BringToFront()
                myfrmTable.BringToFront()
                myfrmTable.LoadThis(myFile.File, strTableName)
            Else
                myfrmTable = New frmTableFields
                myfrmTable.Show()
                myfrmTable.LoadThis(myFile.File, strTableName)
            End If
        End If
    End Sub


    Private Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.F5 Then
            RunButton()
        Else
        End If
    End Sub

    Private Function FileTrim(thisFilePath As String, charLimit As Integer) As String
        Dim i As Integer = 0
        Dim FileCount As Integer = 0
        Dim replacement As String = ""
        Dim FilePath, FilePath2, directoryName, FileName As String
        FilePath = thisFilePath
        FileName = System.IO.Path.GetFileName(thisFilePath)
        FilePath2 = FilePath
        FileCount = FilePath.Count ''No. of Charaters in filepath
        While FileCount > charLimit
            directoryName = System.IO.Path.GetDirectoryName(FilePath)

            ''replacement = directoryName
            FilePath = directoryName
            ''If i = 1 Then
            FilePath2 = FilePath + "\...\" + FileName  ' this will preserve the previous path
            ''End If
            FileCount = FilePath2.Count
            i = i + 1

        End While
        replacement = FilePath2
        Return replacement
    End Function

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Dim strFile As String = ""
        strFile = lblFilePath.Text
        myImport = New frmImport

        myImport.ShowDialog()


        Dim frmCollection = System.Windows.Forms.Application.OpenForms
        If frmCollection.OfType(Of frmImport).Any = False Then
            ClearItems()
            LoadTables(strFile)
            myFile.File = strFile
            FileTrim(strFile, 106)
            tTip.SetToolTip(lblFilePath, myFile.File)
            lblFilePath.Text = strFile
            lblFilePath.Refresh()
        End If

    End Sub

    Private Sub btnOpenForm_Click(sender As Object, e As EventArgs) Handles btnOpenForm.Click
        Dim frmCollection = System.Windows.Forms.Application.OpenForms


        If frmCollection.OfType(Of frmSQL).Any Then
            Me.BringToFront()
            mSQL.BringToFront()
        Else
            mSQL = New frmSQL
            mSQL.ShowDialog()
            mSQL.LoadText(rtxtSelect.Text)
        End If
    End Sub

    Private Sub rtxtSelect_KeyDown(sender As Object, e As KeyEventArgs) Handles rtxtSelect.KeyDown
        If e.Control And (e.KeyCode = Keys.A) Then

            If sender IsNot Nothing Then
                rtxtSelect.SelectAll()
            End If
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub
End Class
