Imports Microsoft.Win32

Public Module MachineLock

    Public Function GetRawMachineGuid() As String

        Try
            ' 1️⃣ Try 64-bit registry (works on 64-bit Windows even from 32-bit DLL)
            Dim guid As String = ReadMachineGuid(RegistryView.Registry64)
            If guid <> "" Then Return guid
        Catch ex As Exception

        End Try

        Try
            ' 2️⃣ Fallback: default registry (for 32-bit Windows)
            dim Guid = ReadMachineGuid(RegistryView.Default)
            If Guid <> "" Then Return Guid
        Catch ex As Exception

        End Try


        ' 3️⃣ Nothing worked
        Return ""

    End Function

    Private Function ReadMachineGuid(view As RegistryView) As String
        Try
            Using baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view)
                Using subKey = baseKey.OpenSubKey("SOFTWARE\Microsoft\Cryptography", False)
                    If subKey Is Nothing Then Return ""
                    Dim value = subKey.GetValue("MachineGuid")


                    If value Is Nothing Then Return ""
                    Return value.ToString().Trim()
                End Using
            End Using
        Catch
            Return ""
        End Try
    End Function

End Module
