Imports System.Data.SqlClient

Public Class Dashboard
    Private Sub Dashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadDashboardStats()
        LoadTopCandidates()
        LoadElectionYears()
    End Sub
    Private Sub LoadDashboardStats()
        Try
            con.Open()

            ' Queries and labels
            Dim queries As New Dictionary(Of String, Label) From {
                {"SELECT COUNT(*) FROM VotersTb", Label9},                       ' Total Voters
                {"SELECT COUNT(*) FROM CandidateTb", Label10},                   ' Total Candidates
                {"SELECT COUNT(DISTINCT VoterID) FROM VoterVotesTb", Label11}    ' Unique Votes Cast
            }

            Dim counts As New List(Of Integer)

            For Each q In queries
                Using cmd As New SqlCommand(q.Key, con)
                    Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    q.Value.Text = count.ToString()
                    counts.Add(count)
                End Using
            Next

            ' Calculate Voter Turnout (%)
            Dim turnout As Double = If(counts(0) > 0, (counts(2) / counts(0)) * 100, 0)
            Label12.Text = Math.Round(turnout, 2).ToString() & " %"

        Catch ex As Exception
            MessageBox.Show("Dashboard Load Error: " & ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        LoadTopCandidates()
    End Sub


    Public Sub LoadTopCandidates()
        Try
            con.Open()
            Dim query As String = "
        SELECT CandidateID, FullName, VotesCount, Platform, YearLevel, Section, Position 
        FROM CandidateTb AS c
        WHERE c.ElectionYear = @ElectionYear
        AND VotesCount = (
            SELECT MAX(VotesCount) 
            FROM CandidateTb 
            WHERE Position = c.Position AND ElectionYear = @ElectionYear
        )
        AND NOT EXISTS (
            SELECT 1 
            FROM CandidateTb 
            WHERE Position = c.Position AND ElectionYear = @ElectionYear AND VotesCount = c.VotesCount AND CandidateID <> c.CandidateID
        )"

            Dim cmd As New SqlCommand(query, con)
            cmd.Parameters.AddWithValue("@ElectionYear", ComboBox1.Text.Trim())

            Dim adapter As New SqlDataAdapter(cmd)
            Dim table As New DataTable()
            adapter.Fill(table)
            DataGridView1.DataSource = table
        Catch ex As Exception
            MessageBox.Show("Error loading top candidates: " & ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Public Sub LoadElectionYears()
        Try
            con.Open()
            Dim query As String = "SELECT DISTINCT ElectionYear FROM CandidateTb ORDER BY ElectionYear DESC"
            Dim cmd As New SqlCommand(query, con)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            ComboBox1.Items.Clear()
            While reader.Read()
                ComboBox1.Items.Add(reader("ElectionYear").ToString())
            End While

            reader.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading election years: " & ex.Message)
        Finally
            con.Close()
        End Try
    End Sub


End Class