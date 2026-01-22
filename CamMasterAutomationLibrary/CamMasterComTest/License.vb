Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports Microsoft.Win32

Public Module License

    Private Const LicensePath As String =
        "C:\ProgramData\VPInformatics\CamMaster\license.key"

    Private Const SECRET As String =
        "VP@2025#CamMaster$Internal"

    Public Sub Validate()

        If Not File.Exists(LicensePath) Then
            Throw New Exception("License file not found.")
        End If

        Dim lines() As String = File.ReadAllLines(LicensePath)
        If lines.Length < 2 Then
            Throw New Exception("Invalid license file.")
        End If

        Dim licensedHash As String = lines(1).Trim().ToLowerInvariant()



        Dim machineGuid As String = MachineLock.GetRawMachineGuid()
        If machineGuid = "" Then
            Throw New Exception("Unable to read machine identifier.")
        End If

        Dim expectedHash As String =
            ComputeSha256(machineGuid & SECRET).ToLowerInvariant()

        If licensedHash <> expectedHash Then
            Throw New Exception("License is not valid for this machine.")
        End If

    End Sub




    Private Function GetMachineGuid() As String
        Using key = Registry.LocalMachine.OpenSubKey(
            "SOFTWARE\Microsoft\Cryptography", False)
            If key IsNot Nothing Then
                Return key.GetValue("MachineGuid", "").ToString()
            End If
        End Using
        Return ""
    End Function

    Private Function ComputeSha256(input As String) As String
        Using sha = SHA256.Create()
            Dim bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(input))
            Dim sb As New StringBuilder()
            For Each b As Byte In bytes
                sb.Append(b.ToString("x2"))
            Next
            Return sb.ToString()
        End Using
    End Function

End Module
