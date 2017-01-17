Imports System.IO
Imports System.Text
Module modGlobal
    Public MAIL_SERVER As String    'Global String to hold Server
    Public MAIL_PORT As Integer     'Global Integer to hold Server Port
    Public MAIL_USER As String      'Global String to hold User Name
    Public MAIL_PASSWORD As String  'Global String to hold User Password
    Public MAIL_USETLS As Boolean   'Global Boolean to hold if SSL is needed
    Public MAIL_TIMEOUT As Long     'Globale Long to hold Server Timeout
    Public MAIL_USEHTML As Boolean
    Public MAIL_ATTACHFILES As String
    Public PROXY_USE As Boolean
    Public PROXY_SERVER As String
    Public PROXY_PORT As Integer
    Public PROXY_UID As String
    Public PROXY_PWD As String
    Public LOG_TO_FILE As Boolean
    Public LOG_NAME As String
    Public LOG_PATH As String
    '''<summary>
    ''' The Get Command will looks for Command Line Arguments, this on will return as string
    ''' the switch will be something like /mystring="this is fun"
    ''' if it is just /mystring then it will return what is set in the sDefault string.
    ''' </summary>
    Public Function GetCommand(ByVal strLookFor As String, ByVal sDefault As String, Optional ByRef DidExist As Boolean = False, Optional ByRef Switch As String = "/") As String
        Dim sAns As String = sDefault
        DidExist = False
        Dim cmdLine() As String = System.Environment.GetCommandLineArgs
        Dim i As Integer = 0
        Dim intCount As Integer = cmdLine.Length
        Dim strValue As String = ""
        If intCount > 1 Then
            For i = 1 To intCount - 1
                strValue = cmdLine(i)
                strValue = Replace(strValue, Switch, "")
                Dim strSplit() As String = Split(strValue, "=")
                Dim intLBound As Integer = LBound(strSplit)
                Dim intUBound As Integer = UBound(strSplit)
                If LCase(strSplit(intLBound)) = LCase(strLookFor) Then
                    If intUBound <> 0 Then
                        sAns = strSplit(intUBound)
                    Else
                        sAns = sDefault
                    End If
                    DidExist = True
                    Exit For
                End If
            Next
        End If
        Return sAns
    End Function
    '''<summary>
    ''' The Get Command will looks for Command Line Arguments, this on will return as long
    ''' the switch will be something like /mylongvalue=92
    ''' if it is just /mylongvalue it will return the lDefault value
    ''' </summary>
    Public Function GetCommand(ByVal strLookFor As String, ByVal lDefault As Long, Optional ByRef DidExist As Boolean = False, Optional ByRef Switch As String = "/") As Long
        Dim lAns As Long = lDefault
        DidExist = False
        Dim cmdLine() As String = System.Environment.GetCommandLineArgs
        Dim i As Integer = 0
        Dim intCount As Integer = cmdLine.Length
        Dim strValue As String = ""
        If intCount > 1 Then
            For i = 1 To intCount - 1
                strValue = cmdLine(i)
                strValue = Replace(strValue, Switch, "")
                Dim strSplit() As String = Split(strValue, "=")
                Dim intLBound As Integer = LBound(strSplit)
                Dim intUBound As Integer = UBound(strSplit)
                If LCase(strSplit(intLBound)) = LCase(strLookFor) Then
                    If intUBound <> 0 Then
                        lAns = strSplit(intUBound)
                    Else
                        lAns = lDefault
                    End If
                    DidExist = True
                    Exit For
                End If
            Next
        End If
        Return lAns
    End Function
    '''<summary>
    ''' The Get Command will looks for Command Line Arguments, this on will return as boolean.
    ''' if the command is /swtich it will return as true since it did exist
    ''' you can also use /switch=false
    ''' </summary>
    Public Function GetCommand(ByVal strLookFor As String, ByVal bDefault As Boolean, Optional ByRef DidExist As Boolean = False, Optional ByRef Switch As String = "/") As Boolean
        Dim bAns As Boolean = bDefault
        DidExist = False
        Dim cmdLine() As String = System.Environment.GetCommandLineArgs
        Dim i As Integer = 0
        Dim intCount As Integer = cmdLine.Length
        Dim strValue As String = ""
        If intCount > 1 Then
            For i = 1 To intCount - 1
                strValue = cmdLine(i)
                strValue = Replace(strValue, Switch, "")
                Dim strSplit() As String = Split(strValue, "=")
                Dim intLBound As Integer = LBound(strSplit)
                Dim intUBound As Integer = UBound(strSplit)
                If LCase(strSplit(intLBound)) = LCase(strLookFor) Then
                    If intUBound <> 0 Then
                        bAns = strSplit(intUBound)
                    Else
                        bAns = True
                    End If
                    DidExist = True
                    Exit For
                End If
            Next
        End If
        Return bAns
    End Function
    Public Sub LogFile(ByVal strPath As String, ByVal strMessage As String)
        Dim SendMessage As String = DateTime.Now & vbTab & strMessage
        Call AppendToFile(strPath, SendMessage)
    End Sub
    Public Sub LogDebugFile(ByVal strPath As String, ByVal strMessage As String)
        Dim SendMessage As String = DateTime.Now & vbTab & strMessage
        Call AppendToFile(strPath, SendMessage)
    End Sub
    Public Function FileExists(ByVal strPath As String)
        Return File.Exists(strPath)
    End Function
    Private Sub CreateFile(ByVal strPath As String)
        If File.Exists(strPath) = False Then
            Dim fs As New FileStream(strPath, FileMode.Append, FileAccess.Write, FileShare.Write)
            fs.Close()
        End If
    End Sub
    Private Sub AppendToFile(ByVal strPath As String, ByVal strNewLine As String)
        If File.Exists(strPath) = False Then Call CreateFile(strPath)
        Dim sw As New StreamWriter(strPath, True, Encoding.ASCII)
        sw.WriteLine(strNewLine)
        sw.Close()
    End Sub
End Module
