Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        startDateTimePicker.Value = DateAdd("d", -1, DateTime.Today)

        endDateTimePicker.Value = DateAdd("d", -1, DateTime.Today)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim access As New MSAccessConnection("PATH_TO_ACCESS_DB.accdb")

        Dim connection As New CMSSupervisorConnection("CMS_SERVER", 2, "Username", "Password")

        connection.Connect()

        Reporting.RunSkillSummary(connection, access, startDateTimePicker.Value.ToString("M/d/yyyy"), endDateTimePicker.Value.ToString("M/d/yyyy"))

        Reporting.RunAgentPerformanceInterval(connection, access, startDateTimePicker.Value.ToString("M/d/yyyy"), endDateTimePicker.Value.ToString("M/d/yyyy"))

        Reporting.RunAgentPerformanceReport(connection, access, startDateTimePicker.Value.ToString("M/d/yyyy"), endDateTimePicker.Value.ToString("M/d/yyyy"))

        Reporting.RunIntervalCallStats(connection, access, startDateTimePicker.Value.ToString("M/d/yyyy"), endDateTimePicker.Value.ToString("M/d/yyyy"))

        Reporting.RunIntervalPerformance(connection, access, startDateTimePicker.Value.ToString("M/d/yyyy"), endDateTimePicker.Value.ToString("M/d/yyyy"))

        connection.Disconnect()

        MsgBox("Reports have completed")

    End Sub

End Class
