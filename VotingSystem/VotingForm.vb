Imports System.Data.SqlClient
Imports System.IO
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Public Class VotingForm


    Private selectedCandidateID As Integer
    Private selectedPosition As String

    Private Sub VotingForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel1.Visible = False


        LoadPositions()

        Me.KeyPreview = True
    End Sub




    Private Sub VotingForm_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        ' Check if the Esc key is pressed
        If e.KeyCode = Keys.Escape Then
            Button3.PerformClick()

        End If
    End Sub
    Private Sub LoadCandidate()
        Try
            Dim query As String = "SELECT CandidateID, FullName, Position, Platform, YearLevel, Section, Profile 
                             FROM CandidateTb 
                             WHERE ElectionYear = YEAR(GETDATE())" ' Filter by current year

            Dim whereClause As String = ""
            Dim cmd As SqlCommand

            If ComboBox6.SelectedIndex <> -1 Then
                whereClause &= " AND Position = @Position"
            End If

            ' Append filter conditions if any
            If whereClause <> "" Then
                query &= whereClause
            End If

            ' Final ordering
            query &= " ORDER BY CandidateID DESC"

            cmd = New SqlCommand(query, con)

            If ComboBox6.SelectedIndex <> -1 Then
                cmd.Parameters.AddWithValue("@Position", ComboBox6.Text)
            End If

            ' Fetch and bind to DataGridView
            Using da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)
                DataGridView1.DataSource = dt

                If dt.Rows.Count = 0 Then
                    MessageBox.Show("No candidate records found for the selected filter(s).", "No Records", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                ' Adjust Profile column to display images
                If DataGridView1.Columns.Contains("Profile") Then
                    Dim profileColumn As DataGridViewImageColumn = CType(DataGridView1.Columns("Profile"), DataGridViewImageColumn)

                    ' Set the image layout to stretch
                    profileColumn.ImageLayout = DataGridViewImageCellLayout.Stretch

                    'Edit size of image
                    DataGridView1.Columns("Profile").Width = 250
                    DataGridView1.RowTemplate.Height = 250
                End If

                ' Resize the images to fit in the DataGridView cells
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Not row.IsNewRow Then
                        If row.Cells("Profile").Value IsNot DBNull.Value Then
                            ' Convert the byte array to Image
                            Dim profileData As Byte() = CType(row.Cells("Profile").Value, Byte())
                            Using ms As New MemoryStream(profileData)
                                Dim originalImage As Image = Image.FromStream(ms)
                                Dim resizedImage As Image = ResizeImage(originalImage, 250, 250)
                                row.Cells("Profile").Value = resizedImage
                            End Using
                        End If
                    End If
                Next

            End Using

        Catch ex As Exception
            MessageBox.Show("Error loading candidates: " & ex.Message)
        End Try
    End Sub

    ' Function to resize the image while maintaining the aspect ratio
    Private Function ResizeImage(ByVal originalImage As Image, ByVal width As Integer, ByVal height As Integer) As Image
        ' Calculate the scaling ratio
        Dim ratioX As Double = width / originalImage.Width
        Dim ratioY As Double = height / originalImage.Height
        Dim ratio As Double = Math.Min(ratioX, ratioY)

        ' Calculate the new size based on the ratio
        Dim newWidth As Integer = Convert.ToInt32(originalImage.Width * ratio)
        Dim newHeight As Integer = Convert.ToInt32(originalImage.Height * ratio)

        ' Create a new resized image
        Dim resizedImage As New Bitmap(originalImage, newWidth, newHeight)
        Return resizedImage
    End Function



    Private Sub LoadPositions()
        Try
            Dim query As String = "SELECT DISTINCT Position FROM CandidateTb"
            Dim cmd As New SqlCommand(query, con)
            con.Open()

            Dim reader As SqlDataReader = cmd.ExecuteReader()
            ComboBox6.Items.Clear()

            While reader.Read()
                ComboBox6.Items.Add(reader("Position").ToString())
            End While

            reader.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading positions: " & ex.Message)
        Finally
            con.Close()
        End Try
    End Sub



    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        LoadCandidate()
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            ' Ensure a valid row is clicked
            If e.RowIndex >= 0 Then
                ' Get CandidateID
                Dim candidateIDCell As Object = DataGridView1.Rows(e.RowIndex).Cells("CandidateID").Value
                If candidateIDCell Is Nothing OrElse IsDBNull(candidateIDCell) Then
                    MessageBox.Show("CandidateID is missing!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                selectedCandidateID = Convert.ToInt32(candidateIDCell)
                selectedPosition = ComboBox6.Text

                ' Show confirmation panel
                Panel1.Visible = True
                Label2.Text = "Are you sure you want to vote for " & DataGridView1.Rows(e.RowIndex).Cells("FullName").Value.ToString() & " as " & selectedPosition & "?"

                ' Get Profile Image from the database
                Dim imageData As Byte() = Nothing
                Dim rowIndex As Integer = e.RowIndex

                ' Ensure "Profile" column exists and contains data
                If DataGridView1.Rows(rowIndex).Cells("Profile").Value IsNot Nothing AndAlso Not IsDBNull(DataGridView1.Rows(rowIndex).Cells("Profile").Value) Then
                    imageData = CType(DataGridView1.Rows(rowIndex).Cells("Profile").Value, Byte())
                End If

                ' Convert Byte Array to Image
                If imageData IsNot Nothing AndAlso imageData.Length > 0 Then
                    Using ms As New MemoryStream(imageData)
                        PictureBox1.Image = Image.FromStream(ms) ' Load image from memory stream
                    End Using
                Else
                    ' If no image found, clear PictureBox or set a default image
                    PictureBox1.Image = Nothing
                    MessageBox.Show("No image available for this candidate.", "No Image", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim voterID As Integer

        ' Check if loggedInUser contains a valid VoterID
        If Integer.TryParse(loggedInUser, voterID) Then
            CastVote(selectedCandidateID, voterID, selectedPosition)
        Else
            MessageBox.Show("Invalid Voter ID! Please log in again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Public Sub CastVote(ByVal candidateID As Integer, ByVal voterID As Integer, ByVal position As String)
        Try
            con.Open()

            ' Check if the voter has already voted for this position
            Dim checkQuery As String = "SELECT COUNT(*) FROM VoterVotesTb WHERE VoterID = @VoterID AND Position = @Position"
            Dim checkCmd As New SqlCommand(checkQuery, con)
            checkCmd.Parameters.AddWithValue("@VoterID", voterID)
            checkCmd.Parameters.AddWithValue("@Position", position)

            Dim voteCount As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

            If voteCount > 0 Then
                MessageBox.Show("You have already voted for this position!", "Voting Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Panel1.Visible = False
                Exit Sub
            End If

            ' Insert vote into VoterVotesTb
            Dim insertQuery As String = "INSERT INTO VoterVotesTb (VoterID, Position, CandidateID, VoteDate) VALUES (@VoterID, @Position, @CandidateID, GETDATE())"
            Dim insertCmd As New SqlCommand(insertQuery, con)
            insertCmd.Parameters.AddWithValue("@VoterID", voterID)
            insertCmd.Parameters.AddWithValue("@Position", position)
            insertCmd.Parameters.AddWithValue("@CandidateID", candidateID)

            insertCmd.ExecuteNonQuery()

            ' Update Candidate Vote Count
            Dim updateQuery As String = "UPDATE CandidateTb SET VotesCount = VotesCount + 1 WHERE CandidateID = @CandidateID"
            Dim updateCmd As New SqlCommand(updateQuery, con)
            updateCmd.Parameters.AddWithValue("@CandidateID", candidateID)

            updateCmd.ExecuteNonQuery()

            ' Update HasVoted in VotersTb
            Dim updateHasVotedQuery As String = "UPDATE VotersTb SET HasVoted = 1 WHERE VoterID = @VoterID"
            Dim updateHasVotedCmd As New SqlCommand(updateHasVotedQuery, con)
            updateHasVotedCmd.Parameters.AddWithValue("@VoterID", voterID)

            updateHasVotedCmd.ExecuteNonQuery()

            MessageBox.Show("Vote successfully cast!", "Voting Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Panel1.Visible = False

        Catch ex As Exception
            MessageBox.Show("Error casting vote: " & ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to log out?", "Confirm Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Dim loginForm As New Form1()
            loginForm.Show()
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Panel1.Visible = False
        DataGridView1.ClearSelection()
    End Sub
End Class