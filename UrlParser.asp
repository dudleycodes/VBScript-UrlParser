<%
''''''''''''''''''''''''''''''''''
'' UrlParser
''      Simple class to parse a URL into its components.
''
'' [Notes]
''      • Creates absolute url from relative url; "dir1/dir2/../dir3" becomes "dir1/dir3"
''      • An expansion idea would be to build this into a UrlBuilder instead of just a parser. It is
''          for this reason the class is build with properties rather than functions.
''
'' [How To Use]
''      Dim Url: Set Url = New UrlParser
''      Url.path = "https://www.server.com/dir1/dir2/../dir3/file.html?i=3&t=5"
''      Response.Write(Url.filename)
''
'' [Exposed Properties (Read Only)]
''      .directories            (array) Returns an array of all directories in the url
''      .directoryCount         (integer) Returns number of directories
''      .directoryString        (string) String containing all directories separated by /
''      .file                   (sting) Filename if any
''      .filename               (string) Alias of .file
''      .fileExtension          (string) File extension (if any)
''      .fullPath               (string) Fully qualitifed url string
''      .host                   (string) Name of host (i.e. server name)
''      .hostname               (string) Alias of .host
''      .pathSeparator          (string) Path separator.  Usually "/"
''      .queries                (dictionary) Collection of url queries:  key => value
''      .queryCount             (integer) Number of query variables
''      .queryString            (string) Query string
''      .scheme                 (string) Scheme used (usually "http" or "https")
''
'' [Exposed Functions]
''      .directory(index)       (string) Returns specified directory name.  Index starts at zero.
''                                  example: In "dir1/dir2/dir3" .directory(1) yields "dir2"
''''''''''''''''''''''''''''''''''
Class UrlParser

    Private m_directoryArray
    Private m_file
    Private m_fileExtension
    Private m_folder
    Private m_host
    Private m_path
    Private m_pathSeparator
    Private m_scheme
    Private m_queryDictionary


    Public Property Get directories
        directories = array()
        If NOT IsNull(m_directoryArray) Then
            directories = m_directoryArray
        End If
    End Property


    Public Function directory(index)
        index = CInt(index)
        If index > UBound(m_directoryArray) OR index < LBound(m_directoryArray) Then
            directory = Null
        Else
            directory = CStr(m_directoryArray(index))
        End If
    End Function


    Public Property Get directoryCount
        directoryCount = 0
        If NOT IsNull(m_directoryArray) Then
            directoryCount = CInt(UBound(m_directoryArray) + 1 - LBound(m_directoryArray))
        End If
    End Property


    Public Property Get directoryString
        directoryString = ""
        If NOT IsNull(m_directoryArray) Then
            directoryString = Join(m_directoryArray, m_pathSeparator)
        End If
    End Property


    Public Property Get file
        file = Null
        If NOT IsEmpty(m_file) Then
            file = CStr(m_file)
        End If
    End Property


    Public Property Get filename
        filename = me.file
    End Property


    Public Property Get fileExtension
        Dim i: i = InStrRev(me.file, ".")
        fileExtension = Null
        If i > 0 Then
            fileExtension = Mid(me.file, i + 1)
        End If
    End Property


    Public Property Get fullPath
        fullPath = CStr(me.scheme & "://" & me.host)

        If LEN(me.directoryString) > 0 Then
            fullPath = fullPath & m_pathSeparator & me.directoryString
        End If

        If LEN(me.file) > 0 Then
            fullPath = fullPath & m_pathSeparator & me.file
        End If

        If LEN(me.queryString) > 0 Then
            fullPath = fullPath & me.queryString
        End If
    End Property


    Public Property Get host
        host = Null
        If NOT IsEmpty(m_host) Then
            host = CStr(m_host)
        End If
    End Property


    Public Property Get hostname
        hostname = me.host
    End Property


    Public Property Get path
        path = Null
        If NOT IsEmpty(m_path) Then
            path = m_path
        End If
    End Property


    Public Property Let path(value)
        Dim i: i = ""
        Dim temp: temp = ""
        Dim temp2: temp2 = ""
        Dim tempPath: tempPath = ""
        Dim regEx: Set regEx = New RegExp

        m_path = CStr(Trim(value))

        'Determine Path Separator
        If InStr(m_path, "/") Then
            m_pathSeparator = "/"
            m_path = Replace(m_path, "\", m_pathSeparator)
        Else
            m_pathSeparator = "\"
            m_path = Replace(m_path, "/", m_pathSeparator)
        End If

        tempPath = m_path
        
        'Determine scheme/host
        i = InStr(tempPath, "://")
        If i > 1 Then
            m_scheme = Left(tempPath, i - 1)
            tempPath = Replace(tempPath, m_scheme & "://", "")
            m_scheme = LCase(m_scheme)
            'Determine host
            m_host = Left(tempPath, InStr(tempPath, m_pathSeparator) - 1)
            tempPath = Replace(tempPath, m_host, "")
        Else
            'Derive scheme
            m_scheme = "http"
            If LCase(Request.ServerVariables("HTTPS")) = "on" Then
                m_scheme = "https"
            End If
            'Derive host
            m_host = Request.ServerVariables("SERVER_NAME")
        End If

        'Add current base path if necessary
        If InStr(tempPath, "../") = 1 Or InStr(tempPath, m_pathSeparator) > 1 Then
            temp = Request.ServerVariables("PATH_INFO")
            tempPath = MID(temp, 1, InStrRev(temp, m_pathSeparator))  & tempPath
        End If

        'Remove any leading path separator
        If InStr(tempPath, m_pathSeparator) = 1 Then
            tempPath = Right(tempPath, LEN(tempPath) - 1)
        End If

        'Resolve any ../ references
        Do
            i = InStr(tempPath, "../")
            If isNull(i) Then
                i = 0
            ElseIf i = 1 Then
                tempPath = MID(tempPath, 4)
            ElseIf i > 0 Then
                'Remove ../ and the preceding directory
                tempPath = Left(tempPath, InStrRev(tempPath, m_pathSeparator, i - 2)) & MID(tempPath, i + 3)
            End If
        Loop While i > 0

        'Process URL Queries
        m_queryDictionary = Null
        i = InStr(1, tempPath, "?")
        If i > 0 Then
            temp = MID(tempPath, i + 1)
            tempPath = Replace(tempPath, "?" & temp, "")
            temp = Split(temp, "&")
            If UBound(temp) <> -1 Then
                Set m_queryDictionary = CreateObject("Scripting.Dictionary")
                For i = LBound(temp) To UBound(temp)
                    temp2 = Split(temp(i), "=")
                    m_queryDictionary.item(temp2(0)) = temp2(1)
                Next
            End If
        End If

        'Process filename
        If InStr(tempPath, m_pathSeparator) Then
            i = InStr(InStrRev(tempPath, m_pathSeparator), tempPath, ".")
        Else
            i = InStr(tempPath, ".")
        End If

        If i Then
            temp = split(tempPath, m_pathSeparator)
            m_file = temp(ubound(temp)) 
            i = Null
            tempPath = Replace(tempPath, m_file, "")
            'Remove trailing path separator
            If Right(tempPath, 1) = m_pathSeparator Then
                tempPath = Left(tempPath, LEN(tempPath) - 1)
            End If
        End If
        
        'Directory Array
        m_directoryArray = Split(tempPath, m_pathSeparator)
        If UBound(m_directoryArray) = -1 Then
            m_directoryArray = Null
        End If
        
        tempPath = Null
    End Property


    Public Property Get pathSeparator
        pathSeparator = CStr(m_pathSeparator)
    End Property


    Public Property Get queries
        Dim temp: Set temp = CreateObject("Scripting.Dictionary")

        If IsNull(m_queryDictionary) Then
            queries = temp
        Else
            queries = m_queryDictionary
        End If
    End Property


    Public Property Get queryCount
        queryCount = 0
        If NOT IsNull(m_queryDictionary) Then
            queryCount = CInt(m_queryDictionary.Count)
        End If
    End Property


    Public Property Get queryString
        Dim element: element = ""

        queryString = ""
        If NOT IsNull(m_queryDictionary) Then
            For Each element in m_queryDictionary
                If queryString = "" Then
                    queryString = "?"
                Else
                    queryString = queryString & "&"
                End If
                queryString = queryString & element & "=" & m_queryDictionary(element)
            Next
        End If
    End Property


    Public Property Get scheme
        scheme = Null
        If NOT IsEmpty(m_scheme) Then
            scheme = CStr(m_scheme)
        End If
    End Property

End Class
%>
