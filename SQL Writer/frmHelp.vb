Imports System.IO

Public Class frmHelp

    Dim myText As New TextVB
    Dim myFile As New FilePath

    Private Sub frmHelp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtHelp.Text = ""

        txtHelp.BackColor = SystemColors.ControlLightLight
        txtHelp.ReadOnly = True
        myText.GiantString()
        LoadTopics()
        AutoScroll = True
    End Sub

    Private Sub LoadTopics()
        lstTopics.Items.Clear()
        lstTopics.Items.Insert(0, "")
        lstTopics.Items.Insert(1, "Select Statements")
        lstTopics.Items.Insert(2, "Insert Statements")
        lstTopics.Items.Insert(3, "Update Statements")
        lstTopics.Items.Insert(4, "Delete Statements")
        lstTopics.Items.Insert(5, "Where Clauses")
        lstTopics.Items.Insert(6, "Creating and Dropping Tables")
        lstTopics.Items.Insert(7, "Advanced")
        lstTopics.Items.Insert(8, "Display all Topics")

    End Sub

    Private Sub SelectTopic()

        Dim intTopic As Integer = 0
        intTopic = lstTopics.SelectedIndex
        Select Case intTopic
            Case 1
                txtHelp.Text = myText.SelectText
            Case 2
                txtHelp.Text = myText.InsertText
            Case 3
                txtHelp.Text = myText.UpdateText
            Case 4
                txtHelp.Text = myText.DeleteText
            Case 5
                txtHelp.Text = myText.WhereText
            Case 6
                txtHelp.Text = myText.CreateText
            Case 7
                txtHelp.Text = myText.AdvanceText
            Case 8
                txtHelp.Text = myText.LargeText
            Case Else
                txtHelp.Text = ""

        End Select

    End Sub

    Private Sub SaveHelpFile()
        Dim intAnswer As Integer = 0
        Dim strFile As String = ""
        Dim strFilePath As String = ""
        Dim strFileName As String = ""
        Dim strFilters As String = ""
        Dim inAnswer As Integer = 0
        Dim SaveFile As New SaveFileDialog
        Dim strDefaultPath As String = ""
        Dim strFileText As String = ""

        If myFile.Help.Length > 4 Then
            strFileName = myFile.Help.Remove(0, InStrRev(myFile.Script, "\"))
            strFilePath = myFile.Help.Replace(strFileName, "")
        End If
        ''
        If strFilePath.Length > 0 Then
            strDefaultPath = strFilePath
        Else
            strDefaultPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End If
        SaveFile.InitialDirectory = strDefaultPath

        SaveFile.FileName = "SQL Help"

        SaveFile.Filter = "Doc File|*.doc|Docx File|*.docx|Rich Text Format File|*.rtf"
        If SaveFile.ShowDialog() = DialogResult.OK Then
            myFile.Help = SaveFile.FileName
            Dim streamWrite As New StreamWriter(SaveFile.FileName)
            strFileText = myText.LargeText

            streamWrite.WriteLine(strFileText)
            streamWrite.Close()


            intAnswer = MsgBox("The script for " & SaveFile.FileName.Remove(0, InStrRev(SaveFile.FileName, "\")) & " was successfully saved." _
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

    Private Sub btnSaveHelp_Click(sender As Object, e As EventArgs) Handles btnSaveHelp.Click
        Try
            SaveHelpFile()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error")
        End Try
    End Sub

    Private Sub lstTopics_DoubleClick(sender As Object, e As EventArgs) Handles lstTopics.DoubleClick
        SelectTopic()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class