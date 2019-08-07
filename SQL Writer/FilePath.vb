Public Class FilePath
    Private filename As String = ""
    Public Property File As String
        Get
            Return filename
        End Get
        Set(ByVal value As String)
            filename = value
        End Set
    End Property

    Private thisFileFrom As String = ""
    Public Property FileFrom As String
        Get
            Return thisFileFrom
        End Get
        Set(ByVal value As String)
            thisFileFrom = value
        End Set
    End Property

    Private scriptFile As String = ""
    Public Property Script As String
        Get
            Return scriptFile
        End Get
        Set(ByVal value As String)
            scriptFile = value
        End Set
    End Property

    Private helpfile As String = ""
    Public Property Help As String
        Get
            Return helpfile
        End Get
        Set(ByVal value As String)
            helpfile = value
        End Set
    End Property
End Class
