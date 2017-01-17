Imports System.Windows.Forms
Imports BurnSoft.Security
Imports System.Configuration
Module modMain
    Dim Message_Body As String
    Dim Message_Subject As String
    Dim Message_To As String
    Dim Message_CC As String
    Dim Message_BCC As String
    Dim Message_From As String
    Dim UPDATEPASSWORD As Boolean
    Dim UPDATEPROXYPWD As Boolean
    ''' <summary>
    ''' Private Function to decrypt the password that is passed via command or in the config file
    ''' </summary>
    ''' <param name="sValue"></param>
    ''' <returns>password in plain text</returns>
    Private Function DCPWD(sValue As String) As String
        Dim sAns As String = ""
        Dim Obj As New Cyhper.oEncrypt
        sAns = Obj.DecryptSHA(sValue)
        Return sAns
    End Function
    ''' <summary>
    ''' Private Function to encyrpt the password that is passed via command or int he config file
    ''' </summary>
    ''' <param name="sValue"></param>
    ''' <returns>password in SHA encryption</returns>
    Private Function ECPWD(sValue As String) As String
        Dim sAns As String = ""
        Dim Obj As New Cyhper.oEncrypt
        sAns = Obj.EncryptSHA(sValue)
        Return sAns
    End Function
    ''' <summary>
    ''' Sub to update the password(s) that are stored in the config file
    ''' Just Pass the new Value, the AppSettings Key Name to update, and 
    ''' the name of the section that it is relating to.
    ''' </summary>
    ''' <param name="sValue"></param>
    ''' <param name="sKey"></param>
    ''' <param name="sSection"></param>
    Private Sub UpdatePasswords(sValue As String, sKey As String, sSection As String)
        Try
            Dim ePWD As String = ECPWD(sValue)
            Dim sPath As String = Application.StartupPath & "\" & My.Application.Info.AssemblyName & ".exe"
            Dim Config As Configuration = System.Configuration.ConfigurationManager.OpenExeConfiguration(sPath)
            Config.AppSettings.Settings.Remove(sKey)
            Config.AppSettings.Settings.Add(sKey, ePWD)
            Config.Save()
            ConfigurationManager.RefreshSection("appSettings")
            Console.WriteLine("The " & sSection & " has been updated!")
        Catch ex As Exception
            Console.WriteLine("Error Occured while saving " & sSection & " setting!")
            Console.WriteLine("Error Details: " & ex.Message.ToString)
            Call ExitApp(1)
        End Try
    End Sub
    ''' <summary>
    ''' Intialize the Global Varaiables from either the app.config or the command parameters
    ''' If the command parameter doesn't exist, then it will default to what is in the app.config
    ''' </summary>
    Sub init()
        Dim DidExist As Boolean = False

        If GetCommand("?", False) Or GetCommand("help", False) Then
            Call ShowHelp()
        End If

        UPDATEPASSWORD = GetCommand("ump", False)
        If UPDATEPASSWORD Then
            Console.Write("Please enter in Password:")
            Dim NewPass As String = Console.ReadLine()
            Call UpdatePasswords(NewPass, "MAIL_PASSWORD", "mail password")
            Call ExitApp()
        End If

        UPDATEPROXYPWD = GetCommand("upp", False)
        If UPDATEPROXYPWD Then
            Console.Write("Please enter in Password:")
            Dim NewPass As String = Console.ReadLine()
            Call UpdatePasswords(NewPass, "PROXY_PWD", "proxy password")
            Call ExitApp()
        End If

        MAIL_SERVER = GetCommand("host", System.Configuration.ConfigurationManager.AppSettings("MAIL_SERVER"))
        If Len(MAIL_SERVER) = 0 Then Call RequirementMissing("/host=[machine]")
        MAIL_PORT = GetCommand("port", CInt(System.Configuration.ConfigurationManager.AppSettings("MAIL_PORT")))
        If Len(MAIL_PORT) = 0 Then Call RequirementMissing("/port=[port]")
        MAIL_TIMEOUT = GetCommand("timeout", CLng(System.Configuration.ConfigurationManager.AppSettings("MAIL_TIMEOUT")))
        If Len(MAIL_TIMEOUT) = 0 Then Call RequirementMissing("/timeout=[timeout]")
        MAIL_USER = GetCommand("user", System.Configuration.ConfigurationManager.AppSettings("MAIL_USER"))

        MAIL_PASSWORD = GetCommand("pwd", DCPWD(System.Configuration.ConfigurationManager.AppSettings("MAIL_PASSWORD")))
        MAIL_USETLS = GetCommand("usetls", CBool(System.Configuration.ConfigurationManager.AppSettings("MAIL_USETLS")))
        MAIL_USEHTML = GetCommand("usehtml", CBool(System.Configuration.ConfigurationManager.AppSettings("MAIL_USEHTML")))

        PROXY_USE = GetCommand("useproxy", CBool(System.Configuration.ConfigurationManager.AppSettings("PROXY_USE")))
        If PROXY_USE Then
            PROXY_PORT = GetCommand("proxyport", CInt(System.Configuration.ConfigurationManager.AppSettings("PROXY_PORT")))
            If Len(PROXY_PORT) = 0 Then Call RequirementMissing("/proxyport=[proxyport]")
            PROXY_SERVER = GetCommand("proxyserver", System.Configuration.ConfigurationManager.AppSettings("PROXY_SERVER"))
            If Len(PROXY_SERVER) = 0 Then Call RequirementMissing("/proxyserver=[proxyserver]")
            PROXY_UID = GetCommand("proxyuser", System.Configuration.ConfigurationManager.AppSettings("PROXY_UID"))
            If Len(PROXY_UID) = 0 Then Call RequirementMissing("/proxyuser=[user id]")
            PROXY_PWD = GetCommand("proxypwd", DCPWD(System.Configuration.ConfigurationManager.AppSettings("PROXY_PWD")))
            If Len(PROXY_PWD) = 0 Then Call RequirementMissing("/proxypwd=[password]")
        End If

        Message_From = GetCommand("from", MAIL_USER)
        If Len(Message_From) = 0 Then Call RequirementMissing("/from=[email]")
        Message_To = GetCommand("to", "", DidExist)
        If Not DidExist Then Call RequirementMissing("/to=[email]")
        Message_CC = GetCommand("cc", "")
        Message_BCC = GetCommand("bcc", "")
        MAIL_ATTACHFILES = GetCommand("attach", "")
        Message_Subject = GetCommand("subject", "", DidExist)
        If Not DidExist Then Call RequirementMissing("/subject=[subject]")
        Message_Body = GetCommand("message", "", DidExist)
        If Not DidExist Then Call RequirementMissing("/message=[message]")

        LOG_TO_FILE = CBool(System.Configuration.ConfigurationManager.AppSettings("LOG_TO_FILE"))
        If LOG_TO_FILE Then
            LOG_NAME = System.Configuration.ConfigurationManager.AppSettings("LOG_NAME")
            LOG_PATH = System.Configuration.ConfigurationManager.AppSettings("LOG_PATH")
            If Len(LOG_PATH) = 0 Then
                LOG_PATH = System.Windows.Forms.Application.StartupPath
            End If
        End If
    End Sub
    ''' <summary>
    ''' Console Message to alert that a required field or configuration is missing
    ''' then it will show the help/about message
    ''' </summary>
    ''' <param name="reqName"></param>
    Sub RequirementMissing(reqName As String)
        Console.WriteLine(reqName & " is missing!  Please review help!")
        Console.WriteLine()
        Call ShowHelp()
    End Sub
    ''' <summary>
    ''' Help/About section to display the switches that can be used for this command
    ''' </summary>
    Sub ShowHelp()
        Console.WriteLine(My.Application.Info.ProductName & " " & My.Application.Info.Copyright)
        Console.WriteLine(String.Format("Version {0}", Application.ProductVersion.ToString))
        Console.WriteLine()
        Console.WriteLine(My.Application.Info.AssemblyName & " /host=[machine] /port=[port] /user=[email] /pwd=[password] /usetsl /to=[email] ")
        Console.WriteLine("/cc=[email] /bcc=[email] /subject=[subject] /message=[message] /help /? [SEE BELOW FOR MORE OPTIONS]" & vbNewLine)
        Console.WriteLine("/host=[machine] - SMTP Server")
        Console.WriteLine("/port=[port] - Port Number to use, usually 25 or 587")
        Console.WriteLine("/timeout=[timeout] - the time out in seconds for the mail to be sent.")
        Console.WriteLine("/user=[email]  - the Login email address for the server, also the from address")
        Console.WriteLine("/pwd=[password] - Password for the User/From email address")
        Console.WriteLine("/ump - Command to only update the stored mail password in the config file.")
        Console.WriteLine("/upp - Command to only update the stored proxy password in the config file.")
        Console.WriteLine("/usetls - Use TLS for connection.")
        Console.WriteLine("/usehtml - enable html format for email.")
        Console.WriteLine("/from=[email] - email from, if the user is the same then this is not needed.")
        Console.WriteLine("/to=[email] - person you want to send to.(comma separated)")
        Console.WriteLine("/cc=[email] - person you want to carbon copy (comma separated)")
        Console.WriteLine("/bcc=[email] - person you want to blind carbon copy (comma separated)")
        Console.WriteLine("/subject=[subject] - The subject of the email, if longer then one word use double qoutes")
        Console.WriteLine("/attach=[filename(s)] - the file(s) that you want to attach to the email.(comma separated)")
        Console.WriteLine("/message=[message] - the body of the message, if longer then one word use double qoutes")
        Console.WriteLine("/useproxy - enable to go through proxy")
        Console.WriteLine("/proxyserver=[proxyserver] - the name or ip address of the proxy")
        Console.WriteLine("/proxyport=[proxyport] - The port that is used for the proxy server.")
        Console.WriteLine("/proxyuser=[user id] - The username that is needed for the proxy")
        Console.WriteLine("/proxypwd=[password] - The password needed for th proxy user listed above..")
        Console.WriteLine("/help, /? - Show help" & vbNewLine)
        Console.WriteLine()
        Console.WriteLine("NOTE: Everything listed above except for the to, cc, bcc, subject, attache and message")
        Console.WriteLine("can also be set in the " & My.Application.Info.AssemblyName & ".exe.config")
        Console.WriteLine()
        Call PressToExit()
    End Sub
    ''' <summary>
    ''' Generic Sub to let the user know that something is done with output and to press
    ''' any key to continue and exit the app
    ''' </summary>
    Sub PressToExit()
        Console.WriteLine()
        Console.WriteLine("Press enter/return key to continue.")
        Console.Read()
        Call ExitApp()
    End Sub
    ''' <summary>
    ''' Closes the application from continuing
    ''' </summary>
    Sub ExitApp(Optional ByVal ExitValue As Integer = 0)
        Application.Exit()
        Environment.Exit(ExitValue)
    End Sub
    ''' <summary>
    ''' Starting Sub
    ''' </summary>
    Sub Main()
        Call init()
        Dim Obj As New BurnSoft.BSMail
        Obj.SendMail(Message_To, Message_From, Message_Subject, Message_Body)
        Obj = Nothing
    End Sub

End Module
