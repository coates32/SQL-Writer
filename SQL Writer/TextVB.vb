Public Class TextVB
    Dim strLarge As String
    Dim strSelect As String
    Dim strInsert As String
    Dim strUpdate As String
    Dim strDelete As String
    Dim strWhere As String
    Dim strCreate As String
    Dim strAdvance As String

    Public ReadOnly Property LargeText As String
        Get
            Return strLarge
        End Get

    End Property

    Public ReadOnly Property SelectText As String
        Get
            Return strSelect
        End Get

    End Property

    Public ReadOnly Property InsertText As String
        Get
            Return strInsert
        End Get
    End Property

    Public ReadOnly Property UpdateText As String
        Get
            Return strUpdate
        End Get
    End Property

    Public ReadOnly Property DeleteText As String
        Get
            Return strDelete
        End Get
    End Property

    Public ReadOnly Property WhereText As String
        Get
            Return strWhere
        End Get
    End Property

    Public ReadOnly Property CreateText As String
        Get
            Return strCreate
        End Get
    End Property

    Public ReadOnly Property AdvanceText As String
        Get
            Return strAdvance
        End Get
    End Property

    Public Sub GiantString()

        Dim strLine As String = "========================================================================="

        strSelect = strLine + vbNewLine + "SELECT statements:

-SELECT statements are used to display data from a table in a database. To write a SELECT statement, write the word 'SELECT', then the fields you want to display (or you can use the '*' to display all of the fields). Inclose the field name with brackets if the field name is a reserve word, like date, or if there are non-letter characters in the field name, like '/' or a <space>. After that, write the table name you'll be pulling the data from.

--Example

SELECT [fieldname1], [fieldname2] FROM TableName *

*This topic assumes that you created a table from the example in 'CREATE and DROP (deleting) tables' topic.
"
        strInsert = strLine + vbNewLine + "INSERT statements (adding records):

-INSERT statements are handy for when you want to add a record to an existing table. To write an INSERT statement, write the words 'INSERT INTO' with a space and type in the table name you want to add a record to. After that, type in the field names inclosed with parenthesis, then type in 'VALUES' with the data values inclosed with parenthesis.

--Example

INSERT INTO TableName ([Fieldname2], [Fieldname3],[Fieldname4]) VALUES ('value2', 2, 1) *

*This topic assumes that you created a table from the example in 'CREATE and DROP (deleting) tables' topic.
"

        strUpdate = strLine + vbNewLine + "Update statements (editing records):

-UPDATE statements are used for when you want to edit a record to an existing table, usually with a WHERE clause. To write an UPDATE statement, write the words UPDATE' with a space and type in the table name you want to update a record to. After that, type in the word 'SET' with the field names and values you. If you don't use a WHERE clause, your UPDATE statement will update ALL of the records.

--Example

Update TableName Set [Fieldname2] = 'value2', [Fieldname3] = value3, [Fieldname4] = 0 *

*This topic assumes that you have created a table from the example in 'CREATE and DROP (deleting) tables' topic.
"

        strDelete = strLine + vbNewLine + "DELETE statements:

-DELETE statements are used for when you want to delete a record to an existing table, usually with a WHERE clause. To write an DELETE statement, write the words 'DELETE' with a space and type in the table name you want to a record to. Like, with UPDATE statement, if you don't use a WHERE clause, your DELETE statement will delete ALL of the records.

--Example

DELETE TableName

"

        strWhere = strLine + vbNewLine + "WHERE clauses:

-The WHERE clause is useful when you want to select, edit, or delete a specific set of records from a table. To use the WHERE clause, type in the word 'WHERE' with the field name and value you want to filter by. Also, you can filter by the field name value even when you don't have the field name in the beginning part of the statement.

--Example(s)
SELECT [fieldname2] FROM TableName WHERE [fieldname1] = value1 **

UPDATE TableName SET [fieldname2] = 'value2', [fieldname3] = 2 WHERE [fieldname1] = 1 **

DELETE TableName WHERE [fieldname1] = 1 AND [fieldname2] = 'value2' **

-One other thing about WHERE clauses. You can use the LIKE statement to filter records by if the condition is similar to what you are looking for. 

--Example
SELECT [fieldname2] FROM TableName WHERE [fieldname1] LIKE '%value1%'

-Notice the apostraphes and the '%' symbols. When the '%' symbol is only in front value1 (or whatever is the value), that means that the statement will look for all of the values in that field that ends with that value. In this example, this select statement will look for all values in fieldname1 that contains 'value1', including 'Thisvalue123' and 'value1'.

**This topic assumes that you have inserted a row from the example in 'INSERT statements (adding records)' topic.
"

        strCreate = strLine + vbNewLine + "CREATE and DROP (deleting) tables:

-You can create a table within a database file with the CREATE TABLE statement, which varies a bit depending on the database software you are using (this guide assumes that you are connecting to an Access database file). See the example below for a basic example of creating a table.

--Example

CREATE TABLE TableName(
	FieldName1 AUTOINCREMENT(1, 1),
	FieldName2 VARCHAR, --OR MEMO 
	FieldName3 integer, 
	FieldName4 bit );

-In the example above, FieldName1 is set to auto-increment by 1. In Access, this field is called an AutoNumber, which a special type of number field that will generate a unique number every newly added field. While not required to have this type of field, it is generally a good idea to have this type of field (and only one of it's type) for when you are creating a new table, which will make updating and delete records much more reliable. In SQL Server, this field is called an Identity field.

-Then we have FieldName2 which is set as a VARCHAR field (holds text data). Unlike AutoNumbers and Integers (more on that later), when you want to update, or filter data in this field, you must surround the value with single quotes, (FieldName2  = 'Your Text Here, Tex') Note that this field in Access can only hold up to 255 characters. If you need to hold more that, use MEMO instead of VARCHAR for the datatype.

-FieldName3 is set as an integer field. Like AutoNumbers, this field only accepts whole numbers (the decimal field can accept numbers like '1.2'). Unlike AutoNumbers, you can have as many of these types of fields as you like since you have to manually put in data instead of it generating a number for you. If you want to know more on different MS Access data types go here: https://msdn.microsoft.com/en-gb/library/windows/desktop/ms714540(v=vs.85).aspx

-FieldName4 is set as a bit field. Think of bit fields like Yes or No, with 1 = 'Yes' and 0 = 'No'. In computer programing, they would call this a boolean. Like AutoNumbers and Integers, this field only accepts whole numbers. Unlike AutoNumbers and Integers, it only accepts 1s and 0s.

-To DROP (or delete) a table, just type in the following below:

--Example

DROP TABLE TableName

"

        strAdvance = strLine + vbNewLine + "Advance stuff:

------------------------------------------------------------------------
-There are some advance aspects to writing in SQL. One of which is the IIF statement.

--Example

IIF(fieldname <> value, [true part], [false part]) as NewName

-In the example above, the IFF statement is made up of 3 parts, broken up by comas. The first is the condition part 'fieldname <> value' where '<>' means not equal. The next part is the true part, which will come true if the condition statement is true. The last part is the false part, which is the very opposite part of the true part. When using the IIF function, you'll need to make sure to enclose all three parts parenthesis and give it a column name as seen in the example

------------------------------------------------------------------------
-JOINs are very powerful when writing SQL statements.

--Example

Select a.fieldnameID, a.fieldname1, b.fieldname3 from table1 as a INNER JOIN table2 as b on a.fieldname1 = b.fieldname1

-In the example above, this will display all of the records in fieldnameID and fieldname1 from the first table and fieldname3 from the second table where the first table's fieldname1 is equal to the second table's fieldname1. Changing the JOIN type from INNER to LEFT will display all of the selected fields' records from the first table and some of the selected fields' records from the second table. 

-Note the 'a.', 'b.', 'as a',  and the 'as b' in the statement above. The part that states the 'as a' and 'as b' are assigning shortcuts to distinguish table1's fieldname1 from table2's fieldname1. Whithout these shortcuts, you would have to write something like table1.fieldname1 to select table1's fieldname1, and that can get messy when doing longer or more complete SQL statements.

------------------------------------------------------------------------
-SELECT INTO statements allows you to copy one datatable from one database (or a file in Access) to another.

--Access Example

Select * INTO [table1] From [table1] IN c:\database.accdb

--SQL Example

Select * INTO NewDatabase.dbo.[table1] From OldDatabase.dbo.[table1]

-Using the Access Example above, you can see how it is similar to a normal SELECT statement with a few differences. For one, the INTO part of the statement will be copying the old table into a newly created datatable that is enclosed with brackets. Also, note what is written after the FROM <table> part of the statement. Since you'll likely be importing datatables from a different database file, you'll need to include the text 'IN <filePath>' where <filePath> will be the full directory for the database file you will be importing your table from.

-You can just use the Import button on the SQL Writer program to simplify this process (it's recommended to do this before trying wrtite the SELECT INTO statement by hand).

-Note: Uncheck the SELECT statement checkbox before using this type of statement.  

-Also you can run multiple non-SELECT statements at the same time by ending each statement with a ';' (semicolon). For an example: INSERT INTO TableName ([Fieldname2], [Fieldname3],[Fieldname4]) VALUES ('value2', 2, 1); INSERT INTO TableName ([Fieldname2], [Fieldname3],[Fieldname4]) VALUES ('value2', 2, 5);

For more information on SQL statements, I suggest that you start by going this link www.w3schools.com/sql/default.asp "

        strLarge = strSelect + vbNewLine + strInsert + vbNewLine + strUpdate + vbNewLine + strDelete

        strLarge += vbNewLine + strWhere + vbNewLine + strCreate + vbNewLine + strAdvance

    End Sub
End Class
