Imports System.Data.SqlClient
Imports System.IO

Public Class VotesRecord
    Private Sub VotesRecord_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadVoteRecords()
        LoadFilterOptions()
    End Sub
    Public Sub LoadVoteRecords()
        Try
            Dim query As String = "SELECT vv.VoteID, v.Fullname AS VoterName, " &
                              "c.FullName AS CandidateName, vv.Position, c.YearLevel, c.Section, vv.VoteDate " &
                              "FROM VoterVotesTb vv " &
                              "INNER JOIN VotersTb v ON vv.VoterID = v.VoterID " &
                              "INNER JOIN CandidateTb c ON vv.CandidateID = c.CandidateID " &
                              "WHERE 1 = 1"

            ' Add filters if CheckBox1 (All) is NOT checked
            If Not CheckBox1.Checked Then
                If ComboBox4.SelectedIndex <> -1 Then
                    query &= " AND c.YearLevel = @YearLevel"
                End If

                If ComboBox5.SelectedIndex <> -1 Then
                    query &= " AND c.Section = @Section"
                End If

                If ComboBox6.SelectedIndex <> -1 Then
                    query &= " AND vv.Position = @Position"
                End If
            End If

            query &= " ORDER BY vv.VoteDate DESC"

            Dim dt As New DataTable()

            Using cmd As New SqlCommand(query, con)
                ' Add parameters only if filters are applied
                If Not CheckBox1.Checked Then
                    If ComboBox4.SelectedIndex <> -1 Then
                        cmd.Parameters.AddWithValue("@YearLevel", ComboBox4.Text)
                    End If
                    If ComboBox5.SelectedIndex <> -1 Then
                        cmd.Parameters.AddWithValue("@Section", ComboBox5.Text)
                    End If
                    If ComboBox6.SelectedIndex <> -1 Then
                        cmd.Parameters.AddWithValue("@Position", ComboBox6.Text)
                    End If
                End If

                con.Open()
                Dim adapter As New SqlDataAdapter(cmd)
                adapter.Fill(dt)
            End Using

            DataGridView1.DataSource = dt

        Catch ex As Exception
            MessageBox.Show("Error loading vote records: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub
    Private Sub LoadFilterOptions()
        Try
            con.Open()

            ' Load YearLevel
            Dim cmdYear As New SqlCommand("SELECT DISTINCT YearLevel FROM CandidateTb", con)
            Dim readerYear = cmdYear.ExecuteReader()
            ComboBox4.Items.Clear()
            While readerYear.Read()
                ComboBox4.Items.Add(readerYear("YearLevel").ToString())
            End While
            readerYear.Close()

            ' Load Section
            Dim cmdSec As New SqlCommand("SELECT DISTINCT Section FROM CandidateTb", con)
            Dim readerSec = cmdSec.ExecuteReader()
            ComboBox5.Items.Clear()
            While readerSec.Read()
                ComboBox5.Items.Add(readerSec("Section").ToString())
            End While
            readerSec.Close()

            ' Load Position
            Dim cmdPos As New SqlCommand("SELECT DISTINCT Position FROM VoterVotesTb", con)
            Dim readerPos = cmdPos.ExecuteReader()
            ComboBox6.Items.Clear()
            While readerPos.Read()
                ComboBox6.Items.Add(readerPos("Position").ToString())
            End While
            readerPos.Close()

        Catch ex As Exception
            MessageBox.Show("Error loading filter options: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        LoadVoteRecords()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        LoadVoteRecords()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        LoadVoteRecords()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        LoadVoteRecords()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        LoadVoteRecords()
    End Sub
End Class