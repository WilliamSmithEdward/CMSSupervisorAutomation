Public Class CMSSupervisorConnection

    Private _serverAddress As String
    Private _acd As Integer
    Private _userName As String
    Private _password As String

    Private _cvsApp As ACSUP.cvsApplication
    Private _cvsConn As ACSCN.cvsConnection
    Private _cvsSrv As ACSUPSRV.cvsServer

    Public Sub New(ByVal serverAddress As String, ByVal acd As Integer, ByVal userName As String, ByVal password As String)

        _serverAddress = serverAddress
        _acd = acd
        _userName = userName
        _password = password

    End Sub

    Public Sub Connect()

        _cvsApp = New ACSUP.cvsApplication
        _cvsConn = New ACSCN.cvsConnection
        _cvsSrv = New ACSUPSRV.cvsServer

        _cvsApp.CreateServer(_userName, "", "", _serverAddress, False, "ENU", _cvsSrv, _cvsConn)

        _cvsConn.Login(_userName, _password, _serverAddress, "ENU")

        _cvsSrv.Reports.ACD = _acd

    End Sub

    Public Sub Disconnect()

        _cvsConn.Logout()
        _cvsConn.Disconnect()

        System.Runtime.InteropServices.Marshal.ReleaseComObject(_cvsApp)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(_cvsConn)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(_cvsSrv)

        _cvsApp = Nothing
        _cvsConn = Nothing
        _cvsSrv = Nothing

    End Sub

    Public Function ExecuteQuery(reportPath As String, reportParams As Dictionary(Of String, String), timeZone As String) As String

        Dim strPath As String = Application.StartupPath
        Dim rep As New ACSREP.cvsReport
        Dim info As ACSCTLG.cvsReportInfo

        If System.IO.File.Exists(strPath & "\export.txt") Then

            System.IO.File.Delete(strPath & "\export.txt")

        End If

        info = _cvsSrv.Reports.Reports(reportPath)

        _cvsSrv.Reports.CreateReport(info, rep)

        If Not timeZone = "" Then

            rep.TimeZone = timeZone

        End If

        For Each kvp As KeyValuePair(Of String, String) In reportParams

            rep.SetProperty(kvp.Key, kvp.Value)

        Next

        rep.ExportData(strPath & "\export.txt", 44, 0, True, True, True)

        rep.Quit()

        If Not _cvsSrv.Interactive Then

            For Each task As Object In _cvsSrv.ActiveTasks

                _cvsSrv.ActiveTasks.Remove(task)

            Next

        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(rep)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(info)

        rep = Nothing
        info = Nothing

        Return My.Computer.FileSystem.ReadAllText(strPath & "\export.txt")

    End Function

End Class
