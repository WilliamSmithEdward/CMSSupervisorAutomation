Public Class Reporting

    Public Shared Sub RunSkillSummary(cms As CMSSupervisorConnection, access As MSAccessConnection, startDate As String, endDate As String)

        Dim sDate As String = startDate

        Do While DateValue(sDate) <= DateValue(endDate)

            If Not access.QueryIntoDataTable("SELECT * FROM [] WHERE [Split/Skill] = '' AND [Date] = #" & String.Format(sDate, "M/d/yyyy") & "#").AsEnumerable.Any(Function(x) x.Field(Of DateTime)("Date") = Convert.ToDateTime(sDate)) Then

                Dim data As String = cms.ExecuteQuery("",
                                New Dictionary(Of String, String) From {{"Splits/Skills", ""}, {"Dates", String.Format(sDate, "M/d/yyyy")}},
                                "")

                Dim lines() As String = data.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                Dim line As String = lines(3)

                Dim fields() As String = line.Split(",")

                Dim sql As String = "
                    INSERT INTO [DATA_Skill_Summary] 
                        (
                            [Date],
                            [Avg Speed Ans],
                            [Avg Aban Time],
                            [Calls Offered],
                            [ACD Calls],
                            [Avg ACD Time],
                            [Avg ACW Time],
                            [ACD Time],
                            [ACW Time],
                            [Hold Time],
                            [Avg Handle Time],
                            [% Within Service Level],
                            [Max Delay],
                            [Calls Answ <30 secs],
                            [Aban Calls],
                            [Dequeued Calls],
                            [Split/Skill]
                        )
                        VALUES ('{0}',{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},'{16}')"

                sql = String.Format(sql, fields(0), fields(1), fields(2), fields(3), fields(4), fields(5), fields(6), fields(7), fields(8), fields(9), fields(10), fields(11), fields(12), fields(13), fields(14), fields(15), "")

                access.ExecuteSql(sql)

                access.ExecuteSql("Execute []")

            End If

            If Not access.QueryIntoDataTable("SELECT * FROM [] WHERE [Split/Skill] = '' AND [Date] = #" & String.Format(sDate, "M/d/yyyy") & "#").AsEnumerable.Any(Function(x) x.Field(Of DateTime)("Date") = Convert.ToDateTime(sDate)) Then

                Dim data As String = cms.ExecuteQuery("",
                                New Dictionary(Of String, String) From {{"Splits/Skills", ""}, {"Dates", String.Format(sDate, "M/d/yyyy")}},
                                "")

                Dim lines() As String = data.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                Dim line As String = lines(3)

                Dim fields() As String = line.Split(",")

                Dim sql As String = "
                    INSERT INTO [DATA_Skill_Summary] 
                        (
                            [Date],
                            [Avg Speed Ans],
                            [Avg Aban Time],
                            [Calls Offered],
                            [ACD Calls],
                            [Avg ACD Time],
                            [Avg ACW Time],
                            [ACD Time],
                            [ACW Time],
                            [Hold Time],
                            [Avg Handle Time],
                            [% Within Service Level],
                            [Max Delay],
                            [Calls Answ <30 secs],
                            [Aban Calls],
                            [Dequeued Calls],
                            [Split/Skill]
                        )
                        VALUES ('{0}',{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},'{16}')"

                sql = String.Format(sql, fields(0), fields(1), fields(2), fields(3), fields(4), fields(5), fields(6), fields(7), fields(8), fields(9), fields(10), fields(11), fields(12), fields(13), fields(14), fields(15), fields(16))

                access.ExecuteSql(sql)

            End If

            sDate = DateAdd("d", 1, sDate)

        Loop

        access.ExecuteSql("Execute []")

    End Sub

    Public Shared Sub RunAgentPerformanceInterval(cms As CMSSupervisorConnection, access As MSAccessConnection, startDate As String, endDate As String)

        Dim sDate As String = startDate

        Do While DateValue(sDate) <= DateValue(endDate)

            If Not access.QueryIntoDataTable("SELECT * FROM [] WHERE [Report_Date] = #" & String.Format(sDate, "M/d/yyyy") & "#").AsEnumerable.Any(Function(x) x.Field(Of DateTime)("Report_Date") = Convert.ToDateTime(sDate)) Then

                Dim data As String = cms.ExecuteQuery("",
                                    New Dictionary(Of String, String) From
                                    {
                                        {"Agent Group", ""},
                                        {"Dates", String.Format(sDate, "M/d/yyyy")},
                                        {"Times", "00:00-23:59"}
                                    },
                                    "US/Pacific")

                Dim lines() As String = data.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                For x = 1 To UBound(lines) - 1

                    Dim fields() As String = lines(x).Split(",")

                    Dim sql As String = "INSERT INTO [] VALUES (#{0}#,'{1}','{2}','{3}','{4}',{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},"

                    sql = sql & "{21},{22},{23},{24},{25},{26},{27},{28},{29},{30},{31},{32},{33},{34},{35})"

                    sql = String.Format(sql, fields(0), fields(1), fields(2), fields(3), fields(4), fields(5), fields(6), fields(7), fields(8), fields(9), fields(10),
                                        fields(11), fields(12), fields(13), fields(14), fields(15), fields(16), fields(17), fields(18), fields(19), fields(20),
                                        fields(21), fields(22), fields(23), fields(24), fields(25), fields(26), fields(27), fields(28), fields(29), fields(30),
                                        fields(31), fields(32), fields(33), fields(34), fields(35))

                    access.ExecuteSql(sql)

                Next

            End If

            sDate = DateAdd("d", 1, sDate)

        Loop

        access.ExecuteSql("Execute []")

    End Sub

    Public Shared Sub RunAgentPerformanceReport(cms As CMSSupervisorConnection, access As MSAccessConnection, startDate As String, endDate As String)

        Dim sDate As String = startDate

        Do While DateValue(sDate) <= DateValue(endDate)

            If Not access.QueryIntoDataTable("SELECT * FROM [] WHERE [Report Date] = #" & String.Format(sDate, "M/d/yyyy") & "#").AsEnumerable.Any(Function(x) x.Field(Of DateTime)("Report Date") = Convert.ToDateTime(sDate)) Then

                Dim data As String = cms.ExecuteQuery("",
                                    New Dictionary(Of String, String) From
                                    {
                                        {"Agent Group", ""},
                                        {"Date", String.Format(sDate, "M/d/yyyy")}
                                    },
                                    "")

                Dim lines() As String = data.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                For x = 4 To UBound(lines)

                    Dim fields() As String = lines(x).Split(",")

                    Dim sql As String = "INSERT INTO [] VALUES (#{0}#,'{1}',{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15})"

                    sql = String.Format(sql, String.Format(sDate, "M/d/yyyy"), fields(0) & ", " & Right(fields(1), Len(fields(1)) - 1), fields(2), fields(3), fields(4), fields(5), fields(6), fields(7), fields(8), fields(9), fields(10),
                                        fields(11), fields(12), fields(13), fields(14), fields(15))

                    access.ExecuteSql(sql)

                Next

            End If

            sDate = DateAdd("d", 1, sDate)

        Loop

    End Sub

    Public Shared Sub RunIntervalCallStats(cms As CMSSupervisorConnection, access As MSAccessConnection, startDate As String, endDate As String)

        Dim sDate As String = startDate

        Do While DateValue(sDate) <= DateValue(endDate)

            If Not access.QueryIntoDataTable("SELECT * FROM [] WHERE [Report_Date] = #" & String.Format(sDate, "M/d/yyyy") & "#").AsEnumerable.Any(Function(x) x.Field(Of DateTime)("Report_Date") = Convert.ToDateTime(sDate)) Then

                Dim data As String = cms.ExecuteQuery("",
                                    New Dictionary(Of String, String) From
                                    {
                                        {"Splits/Skills", ""},
                                        {"Dates", String.Format(sDate, "M/d/yyyy")},
                                        {"Times", "00:00-23:59"}
                                    },
                                    "US/Pacific")

                Dim lines() As String = data.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                For x = 4 To UBound(lines)

                    Dim fields() As String = lines(x).Split(",")

                    Dim sql As String = "INSERT INTO [] VALUES (#{0}#,'{1}',{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20})"

                    sql = String.Format(sql, String.Format(sDate, "M/d/yyyy"), fields(0), fields(1), fields(2), fields(3), fields(4), fields(5), fields(6), fields(7), fields(8), fields(9), fields(10),
                                        fields(11), fields(12), fields(13), fields(14), fields(15), fields(16), fields(17), fields(18), fields(19))

                    access.ExecuteSql(sql)

                Next

            End If

            sDate = DateAdd("d", 1, sDate)

        Loop

        access.ExecuteSql("Execute []")

    End Sub

    Public Shared Sub RunIntervalPerformance(cms As CMSSupervisorConnection, access As MSAccessConnection, startDate As String, endDate As String)

        Dim sDate As String = startDate

        Do While DateValue(sDate) <= DateValue(endDate)

            If Not access.QueryIntoDataTable("SELECT * FROM [] WHERE [Report_Date] = #" & String.Format(sDate, "M/d/yyyy") & "#").AsEnumerable.Any(Function(x) x.Field(Of DateTime)("Report_Date") = Convert.ToDateTime(sDate)) Then

                Dim data As String = cms.ExecuteQuery("",
                                    New Dictionary(Of String, String) From
                                    {
                                        {"Agent Group", ""},
                                        {"Dates", String.Format(sDate, "M/d/yyyy")},
                                        {"Times", "00:00-23:59"}
                                    },
                                    "US/Pacific")

                Dim lines() As String = data.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                For x = 5 To UBound(lines)

                    Dim fields() As String = lines(x).Split(",")

                    Dim sql As String = "INSERT INTO [] VALUES (#{0}#,'{1}',{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12})"

                    sql = String.Format(sql, String.Format(sDate, "M/d/yyyy"), fields(0) & " " & fields(1) & " " & fields(2), fields(3), fields(4), fields(5), fields(6), fields(7), fields(8), fields(9), fields(10),
                                        fields(11), fields(12), fields(13))

                    access.ExecuteSql(sql)

                Next

            End If

            sDate = DateAdd("d", 1, sDate)

        Loop

        access.ExecuteSql("Execute []")

    End Sub

End Class
