Imports Microsoft.Win32
Imports System.IO

Public Class Main_Screen

    '[HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001]
    '"Account Name"="172.16.1.254"
    '"Connection Type"=dword:00000003
    '"IMAP Server"="172.16.1.254"
    '"IMAP User Name"=""
    '"IMAP Use Sicily"=dword:00000000
    '"IMAP Prompt for Password"=dword:00000001
    '"SMTP Server"="172.16.1.254"
    '"SMTP Display Name"="Tsiba Student"
    '"SMTP Email Address"=""

    '[HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001]
    '"Account Name"="172.16.1.254"
    '"Connection Type"=dword:00000003
    '"IMAP Server"="172.16.1.254"
    '"IMAP User Name"="t0532"
    '"IMAP Prompt for Password"=dword:00000001
    '"SMTP Server"="172.16.1.254"
    '"SMTP Display Name"="Tsiba Student"
    '"SMTP Email Address"="t0532@tsiba.org.za"
    '"IMAP Polling"=dword:00000001
    '"IMAP Dirty"=dword:00000000

    Private Sub Error_Handler(ByVal except As Exception, ByVal message As String)
        Try
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\TSiBA_Error_Log.txt").Replace("\\", "\"), True)
            filewriter.WriteLine(Format(Now(), "ddMMyyyy HH:mm:ss") & ": " & message & vbCrLf & except.ToString & vbCrLf)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            MsgBox("Critical error trapped in Error Handler:" & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.Critical, "Critical Error Trapped")
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim userName As String = My.User.Name
            If userName.LastIndexOf("\") <> -1 Then
                userName = userName.Remove(0, userName.LastIndexOf("\") + 1)
            End If

            Dim exists As Boolean = False
            Try
                If My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Internet Account Manager\Accounts\00000001") IsNot Nothing Then
                    exists = True
                Else
                    Dim newKey As RegistryKey
                    newKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Internet Account Manager\Accounts\00000001")
                End If
            Finally
                My.Computer.Registry.CurrentUser.Close()
            End Try

            Dim server, email, displayName As String
            server = ""
            email = ""
            displayName = ""


            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\TSiBA_Config.txt").Replace("\\", "\")) = True Then
                Dim filereader As StreamReader = New StreamReader((Application.StartupPath & "\TSiBA_Config.txt").Replace("\\", "\"))
                Dim lineread As String = ""
                While filereader.Peek() <> -1
                    lineread = filereader.ReadLine
                    lineread = lineread.Replace(" = ", "=")
                    If lineread.StartsWith("server=", True, System.Globalization.CultureInfo.CurrentCulture) = True Then
                        server = lineread.Remove(0, 7)
                    End If
                    If lineread.StartsWith("email=", True, System.Globalization.CultureInfo.CurrentCulture) = True Then
                        email = lineread.Remove(0, 6)
                    End If
                    If lineread.StartsWith("displayName=", True, System.Globalization.CultureInfo.CurrentCulture) = True Then
                        displayName = lineread.Remove(0, 12)
                    End If
                End While
                filereader.Close()
            End If

            If server.Length < 1 Then
                server = "172.16.1.254"
            End If
            If email.Length < 1 Then
                email = "@tsiba.org.za"
            End If
            If displayName.Length < 1 Then
                displayName = "Tsiba Student"
            End If

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001", "Account Name", server, RegistryValueKind.String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001", "Connection Type", "00000003", RegistryValueKind.DWord)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001", "IMAP Server", server, RegistryValueKind.String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001", "IMAP User Name", userName, RegistryValueKind.String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001", "IMAP Prompt for Password", "00000001", RegistryValueKind.DWord)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001", "SMTP Server", server, RegistryValueKind.String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001", "SMTP Display Name", displayName, RegistryValueKind.String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001", "SMTP Email Address", userName & email, RegistryValueKind.String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001", "IMAP Polling", "00000001", RegistryValueKind.DWord)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\00000001", "IMAP Dirty", "00000000", RegistryValueKind.DWord)

        Catch ex As Exception
            Error_Handler(ex, "Application Load")
        End Try
        Me.Close()
    End Sub
End Class
