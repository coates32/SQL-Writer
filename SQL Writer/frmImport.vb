Option Explicit On
Option Strict On
Imports System.Data.OleDb

Public Class frmImport
    Dim myFile As New FilePath
    Dim myText As New TextVB
    Dim fileDes As String = ""
    Dim fileOrigin As String = ""
    Dim myTables, deleteMyTables As New List(Of String)

    Private Sub frmImport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DataGridView1.Columns.Add("tbl_No", "Table Number")
        DataGridView1.Columns.Add("tbl_Name", "Table Name")

        Dim chk As New DataGridViewCheckBoxColumn()
        DataGridView1.Columns.Add(chk)
        chk.HeaderText = "Add to Import"
        chk.Name = "Import"
        chk.FalseValue = CheckState.Unchecked
        chk.TrueValue = CheckState.Checked


        chk.ReadOnly = False
    End Sub

    Private Sub populateGrid(thisFile As String)

        Dim userTables As DataTable = Nothing
        Dim con1 As New OleDbConnection

        Dim restrictions() As String = New String(3) {}
        Try
            restrictions(3) = "Table"

            con1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & thisFile
            con1.Open()
            userTables = con1.GetSchema("Tables", restrictions)

            con1.Close()

            For i As Integer = 0 To userTables.Rows.Count - 1 Step i + 1

                Dim thisTable As String
                thisTable = userTables.Rows(i)(2).ToString

                DataGridView1.Rows.Add(i, thisTable)

                DataGridView1.Item(0, i).ReadOnly = True
                DataGridView1.Item(1, i).ReadOnly = True

            Next
        Catch ex As Exception

            con1.Close()

        End Try
        If con1.State.ToString.ToLower.Equals("open") Then
            con1.Close()
        End If
    End Sub

    Private Sub btnOrigin_Click(sender As Object, e As EventArgs) Handles btnOrigin.Click
        Try
            'Opens Access DB file
            OpenDB()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error")
        End Try
    End Sub

    Private Sub OpenDB2() ''Import To

        Dim strFile As String = ""
        Dim strFilePath As String = ""
        Dim strFileName As String = ""
        Dim strFilters As String = ""
        Dim inAnswer As Integer = 0
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim strDefaultPath As String = ""
        If fileDes.Length < 1 Then
            fileDes = myFile.File
        End If
        If fileDes.Length > 4 Then
            strFileName = fileDes.Remove(0, InStrRev(fileDes, "\"))
            strFilePath = fileDes.Replace(strFileName, "")
        End If
        ''
        If strFilePath.Length > 0 Then
            strDefaultPath = strFilePath
        Else
            strDefaultPath = "Z:\K drive (Extra)\Programming Projects\SQL Writer\SQL Writer\bin\Debug"
        End If
        OpenFileDialog1.InitialDirectory = strDefaultPath

        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Access 2007 file (*.Accdb)|*.Accdb|Access 2003 file (*.mdb)|*.mdb|All Access filetypes|*.Accdb; *.mdb"
        OpenFileDialog1.Multiselect = False

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            strFile = OpenFileDialog1.FileName

            fileDes = strFile
            strFile = FileTrim(strFile, 87)
            lblDestination.Text = strFile
            tTip.SetToolTip(lblDestination, fileDes)

            lblDestination.Refresh()
        End If
    End Sub

    Private Sub OpenDB() ''Import From

        Dim strFile As String = ""
        Dim strFilePath As String = ""
        Dim strFileName As String = ""
        Dim strFilters As String = ""
        Dim inAnswer As Integer = 0
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim strDefaultPath As String = ""
        If fileOrigin.Length > 4 Then
            strFileName = fileOrigin.Remove(0, InStrRev(fileOrigin, "\"))
            strFilePath = fileOrigin.Replace(strFileName, "")
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
            ''ClearItems()
            populateGrid(strFile)
            fileOrigin = strFile
            strFile = FileTrim(strFile, 87)
            lblOrigin.Text = strFile
            tTip.SetToolTip(lblOrigin, fileOrigin)
            lblOrigin.Refresh()
        End If
    End Sub
    Private Sub PopCheckboxes()
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim cell As New DataGridViewCheckBoxCell()
            DataGridView1(2, row.Index) = cell
        Next
    End Sub
    Private Sub ClearItems()
        fileDes = ""
        fileOrigin = ""

        If DataGridView1.RowCount > 0 Then
            DataGridView1.Rows.Clear()
        End If
        myTables.Clear()
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        myTables.Clear()
        If DataGridView1.RowCount > 0 Then
            GetChecked()
            CheckTable()
        Else
            MsgBox("No tables in file.")
        End If
    End Sub

    Private Sub GetChecked()
        For i As Integer = 0 To DataGridView1.RowCount - 1
            Dim tableName As String

            If Convert.ToInt32(DataGridView1.Item(2, i).Value) = 1 Then
                tableName = DataGridView1.Item(1, i).Value.ToString
                myTables.Add(tableName)
            End If


        Next

    End Sub
    Private Sub CheckTable()

        Dim tableNames As String = ""
        Dim con1 As New OleDbConnection
        Dim i As Integer = 0
        con1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileDes

        Dim sql As String = ""


        Dim userTables As DataTable = Nothing

        Dim restrictions() As String = New String(3) {}
        Try
            restrictions(3) = "Table"

            con1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileDes
            con1.Open()
            userTables = con1.GetSchema("Tables", restrictions)

            con1.Close()

            For i = 0 To userTables.Rows.Count - 1 Step i + 1
                For index As Integer = 0 To myTables.Count - 1 Step 1
                    If myTables(index).ToString.Equals(userTables.Rows(i)(2).ToString) Then
                        tableNames += userTables.Rows(i)(2).ToString + ","
                        deleteMyTables.Add(userTables.Rows(i)(2).ToString)
                    End If
                Next
            Next

        Catch ex As Exception

            con1.Close()

        End Try
        If con1.State.ToString.ToLower.Equals("open") Then
            con1.Close()
        End If
        ''end loop

        If tableNames.Equals("") Then
            ImportIntoDB()

        Else
            tableNames = "-" & tableNames
            tableNames += "!"
            tableNames = tableNames.Replace(",!", "")
            tableNames = tableNames.Replace(",", ", ")
            Dim tableNames2 As String = tableNames
            tableNames = tableNames.Replace(",", vbCrLf & "-")
            ''vbCrLf 
            Dim message As String
            Dim answer As Integer = 0
            Dim thisFileName = System.IO.Path.GetFileName(fileDes)
            message = "The following table(s) are currently in the destination file '" & thisFileName & "': " & tableNames & "."
            message += vbCrLf & "Do you want to overwrite the existing table(s)?"
            answer = MsgBox(message, MsgBoxStyle.YesNo, "Information")

            If answer = vbYes Then

                If deleteMyTables.Count > 0 Then
                    DeleteTables()
                End If

                ImportIntoDB()
            End If

        End If
    End Sub
    Private Sub DeleteTables()
        '' Create documentation for table import
        Dim i As Integer = 0
        Dim sql As String ' SQL Statement
        Dim con1 As New OleDbConnection

        con1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileDes
        If deleteMyTables.Count > 0 Then
            con1.Open()
            Try

                For i = 0 To deleteMyTables.Count - 1

                    Dim strOriginPath = "'" & fileOrigin & "'"

                    sql = "DROP Table [" & deleteMyTables(i).ToString() & "] "

                    Dim cmd As New OleDbCommand(sql, con1)
                    cmd.ExecuteNonQuery()
                Next
                If con1.State.ToString.ToLower.Equals("open") Then
                    con1.Close()
                End If

            Catch ex As Exception
                MsgBox("Error on line " & (i + 1) & ". " & ex.Message, MsgBoxStyle.OkOnly, "Error")
                If con1.State.ToString.ToLower.Equals("open") Then
                    con1.Close()
                End If
            End Try
        End If
        deleteMyTables.Clear()
    End Sub
    Private Sub ImportIntoDB()
        '' Create documentation for table import
        Dim i As Integer = 0

        Dim con1 As New OleDbConnection

        con1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileDes
        If myTables.Count > 0 Then
            con1.Open()
            Dim sql As String ' SQL Statement

            Try
                For i = 0 To myTables.Count - 1

                    Dim strOriginPath = "'" & fileOrigin & "'"

                    sql = "Select * INTO [" & myTables(i).ToString() & "] From [" _
                        & myTables(i).ToString() & "] In " & strOriginPath

                    Dim cmd As New OleDbCommand(sql, con1)
                    cmd.ExecuteNonQuery()
                Next
                If con1.State.ToString.ToLower.Equals("open") Then
                    con1.Close()
                End If
                MsgBox("Import Complete.", MsgBoxStyle.OkOnly, "Table Import")
            Catch ex As Exception
                MsgBox("Error on line " & (i + 1) & ". " & ex.Message, MsgBoxStyle.OkOnly, "Error")
                If con1.State.ToString.ToLower.Equals("open") Then
                    con1.Close()
                End If
            End Try
        Else
            MsgBox("You didn't select any tables.")
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
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
    Private Sub btnDestination_Click(sender As Object, e As EventArgs) Handles btnDestination.Click
        Try
            'Opens Access DB file
            OpenDB2()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error")
        End Try
    End Sub
End Class