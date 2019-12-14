Option Explicit On
Option Strict On

Imports System.Data.OleDb

'' Reference Microsoft ActiveX Data Objects 6.1
Public Class frmTableFields
    Private Sub frmTableFields_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Font = New Font("Arial", 12, FontStyle.Regular)
        DataGridView1.AutoSizeColumnsMode =
        DataGridViewAutoSizeColumnsMode.AllCells
    End Sub

    Public Sub LoadThis(ByVal Optional thisFile As String = "", ByVal Optional tableName As String = "")
        lblTableName.Text = "Columns in table name: " + tableName
        lblTableName.Refresh()
        LoadColumns(thisFile, tableName)
    End Sub

    Private Sub LoadColumnsNew(ByVal thisFile As String, ByVal tableName As String)
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        Dim row1, autoIndex As Integer

        Dim autoColumn As String = ""
        row1 = 0
        autoIndex = -1
        Dim con1 As New OleDbConnection

        Dim cn = New ADODB.Connection()
        Dim rs = New ADODB.Recordset()
        Dim cnStr As String = ""
        Dim intFields As Integer = 0

        con1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & thisFile

        Dim filterValues = {Nothing, Nothing, tableName, Nothing}

        Try

            Dim i As Integer = 0
            Dim strColumn As String = ""
            Dim strDataType As String = ""
            Dim strAuto As String = ""
            DataGridView1.Columns.Add("ColumnName", "Column Name")
            DataGridView1.Columns.Add("DataType", "Data Type")

            cnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & thisFile
            Dim ssql As String = "Select * From " & tableName

            cn.Open(cnStr, Nothing, Nothing, 0)

            rs.Open(ssql, cn, ADODB.CursorTypeEnum.adOpenKeyset,
                ADODB.LockTypeEnum.adLockOptimistic, -1)
            Dim fld As ADODB.Field

            For Each fld In rs.Fields
                strColumn = fld.Name.ToString
                intFields = Convert.ToInt32(CInt(fld.Type).ToString())
                strDataType = ConvertDataTypeValue(CInt(intFields)).ToString
                DataGridView1.Rows.Add(strColumn, strDataType)
            Next




            autoColumn = GetAutoNumber(thisFile, tableName)

            If autoColumn.Equals("") Then

            Else
                For row = 0 To DataGridView1.RowCount - 0 Step 1
                    If DataGridView1.Item(0, row1).Value.Equals(autoColumn) Then
                        autoIndex = row1
                        Exit For
                    End If
                    row1 += 1
                Next

            End If
            If autoIndex <> -1 Then
                DataGridView1.Item(1, autoIndex).Value = "AutoNumber"
                DataGridView1.Refresh()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error")
        End Try
    End Sub

    Private Sub LoadColumns(ByVal thisFile As String, ByVal tableName As String)
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        Dim row1, autoIndex As Integer

        Dim autoColumn As String = ""
        row1 = 0
        autoIndex = -1
        Dim con1 As New OleDbConnection

        Dim cn = New ADODB.Connection()
        Dim rs = New ADODB.Recordset()
        Dim cnStr As String = ""
        Dim strFields As String = ""

        con1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & thisFile

        Dim filterValues = {Nothing, Nothing, tableName, Nothing}

        Try
            Using conn = New OleDbConnection(con1.ConnectionString)
                conn.Open()
                Dim columns = conn.GetSchema("Columns", filterValues)
                Dim i As Integer = 0
                Dim strColumn As String = ""
                Dim strDataType As String = ""
                Dim strAuto As String = ""
                DataGridView1.Columns.Add("ColumnName", "Column Name")
                DataGridView1.Columns.Add("DataType", "Data Type")
                For Each row As DataRow In columns.Rows
                    strColumn = row("column_name").ToString
                    ''Dim em = columns.Rows(0)("CHARACTER_MAXIMUM_LENGTH").ToString

                    ''Even when a field is set as a MEMO field it may display as "Varchar" without this IF statement
                    If CheckMemo(thisFile, tableName, strColumn) = 203 Then
                        strDataType = ConvertDataTypeValue(203).ToString()
                    Else
                        strDataType = ConvertDataTypeValue(CInt(row("data_type"))).ToString '' Gets the DataType value and converts it to a Value Naume

                    End If

                    If strDataType.ToLower.Equals("varchar") Then
                        Dim colWidth As Integer = CInt(row("CHARACTER_MAXIMUM_LENGTH").ToString)
                        strDataType += "(" & colWidth.ToString & ")"
                    End If
                    DataGridView1.Rows.Add(strColumn, strDataType)
                    DataGridView1.Refresh()
                    ''i += 1
                Next
                conn.Close()
            End Using
            DataGridView1.Refresh()
            autoColumn = GetAutoNumber(thisFile, tableName)

            If autoColumn.Equals("") Then

            Else
                For row = 0 To DataGridView1.RowCount - 0 Step 1
                    If DataGridView1.Item(0, row1).Value.Equals(autoColumn) Then
                        autoIndex = row1
                        Exit For
                    End If
                    row1 += 1
                Next

            End If
            If autoIndex <> -1 Then
                DataGridView1.Item(1, autoIndex).Value = "AutoNumber"
                DataGridView1.Refresh()
            End If
            DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error")
        End Try
    End Sub

    ''Checks to see if field is a Memo (holds more than 255 characters) field or a Varchar field
    Private Function CheckMemo(ByVal thisFile As String, ByVal tableName As String, ByVal strColumn As String) As Integer

        Dim row1, autoIndex As Integer

        Dim autoColumn As String = ""
        row1 = 0
        autoIndex = -1
        Dim con1 As New OleDbConnection

        Dim cn = New ADODB.Connection()
        Dim rs = New ADODB.Recordset()
        Dim cnStr As String = ""
        Dim intFields As Integer = 0

        con1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & thisFile

        Dim filterValues = {Nothing, Nothing, tableName, Nothing}

        Try

            Dim i As Integer = 0

            Dim strDataType As String = ""
            Dim strAuto As String = ""

            cnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & thisFile
            Dim ssql As String = "Select * From " & tableName

            cn.Open(cnStr, Nothing, Nothing, 0)

            rs.Open(ssql, cn, ADODB.CursorTypeEnum.adOpenKeyset,
                ADODB.LockTypeEnum.adLockOptimistic, -1)
            Dim fld As ADODB.Field

            For Each fld In rs.Fields
                If fld.Name = strColumn Then
                    intFields = Convert.ToInt32(CInt(fld.Type).ToString())
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error")
        End Try
        Return intFields
    End Function

    Private Function GetAutoNumber(ByVal thisFile As String, ByVal tableName As String) As String
        Dim sql As String = "SELECT * from " & tableName
        Dim dt As New DataTable()
        Dim con As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & thisFile
        Dim strColumn As String = ""
        Dim dataAdapter As New OleDb.OleDbDataAdapter(sql, con)
        dataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
        dataAdapter.Fill(dt)
        dataAdapter.Dispose()

        For Each column As DataColumn In dt.Columns
            Dim i As String = column.DataType.ToString
            If column.AutoIncrement = True Then
                strColumn = column.ColumnName
                Return strColumn
            End If
        Next
        Return ""
    End Function

    Private Function ConvertDataTypeValue(ByVal oleDbDataType As Integer) As String

        Dim strDataType As String = ""
        Select Case oleDbDataType

            Case OleDbType.LongVarChar
                strDataType = "Varchar"
            Case OleDbType.BigInt
                strDataType = "Int"  '' In Jet this is 32 bit while bigint is 64 bits
            Case OleDbType.LongVarBinary, OleDbType.Binary
                strDataType = "Binary"
            Case OleDbType.Boolean
                strDataType = "Bit"
            Case OleDbType.Char
                strDataType = "Char"
            Case OleDbType.Currency
                strDataType = "Decimal"
            Case OleDbType.DBTimeStamp, OleDbType.DBDate, OleDbType.Date
                strDataType = "Datetime"
            Case OleDbType.Numeric, OleDbType.Decimal
                strDataType = "Decimal"
            Case OleDbType.Double
                strDataType = "Double"
            Case OleDbType.Integer
                strDataType = "Int"
            Case OleDbType.Single
                strDataType = "Single"
            Case OleDbType.SmallInt
                strDataType = "Smallint"
            Case OleDbType.TinyInt
                strDataType = "Smallint"  '' Signed byte not handled by jet so we need 16 bits
            Case OleDbType.UnsignedTinyInt
                strDataType = "Byte"
            Case OleDbType.VarBinary
                strDataType = "Varbinary"
            Case OleDbType.VarChar
                strDataType = "Varchar"
            Case OleDbType.WChar
                strDataType = "Varchar"
            Case OleDbType.LongVarWChar
                strDataType = "Memo"
        End Select
        Return strDataType
    End Function

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim strDataType As String
        Try
            Dim index As Integer
            index = DataGridView1.CurrentRow.Index

            strDataType = DataGridView1.Item(0, index).Value.ToString
            frmMain.rtxtSelect.Text = frmMain.rtxtSelect.Text + strDataType + " "
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub


End Class