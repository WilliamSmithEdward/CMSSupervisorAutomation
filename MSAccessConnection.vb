Public Class MSAccessConnection

    Dim _dbFilePath As String

    Public Sub New(dbFilePath As String)

        _dbFilePath = dbFilePath

    End Sub

    Public Function QueryIntoDataTable(sql As String) As DataTable

        Using cnn As New OleDb.OleDbConnection

            cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _dbFilePath & ";Persist Security Info=False;"

            Using cmd As New OleDb.OleDbCommand(sql, cnn)

                Using da As New OleDb.OleDbDataAdapter(cmd)

                    Using dt As New DataTable

                        da.Fill(dt)

                        Return dt

                    End Using

                End Using

            End Using

        End Using

    End Function

    Public Sub ExecuteSql(sql As String)

        Using cnn As New OleDb.OleDbConnection

            cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _dbFilePath & ";Persist Security Info=False;"

            cnn.Open()

            Using cmd As New OleDb.OleDbCommand(sql, cnn)

                cmd.ExecuteNonQuery()

            End Using

        End Using

    End Sub

End Class
