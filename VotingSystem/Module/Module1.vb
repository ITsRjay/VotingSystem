Imports System.Data.SqlClient
Imports System.Drawing.Imaging
Imports System.IO

Module Module1
    Public con As New SqlConnection("Data Source=MSOO\SQLEXPRESS;Initial Catalog=Voting;Integrated Security=True")
    Public loggedInUser As String
    Public loggedInRole As String

    Public Function ImageToByteArray(image As Image) As Byte()
        Using ms As New System.IO.MemoryStream()
            ' Clone the image to avoid issues with GDI+ stream locks
            Using cloneImage As New Bitmap(image)
                cloneImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            End Using
            Return ms.ToArray()
        End Using
    End Function




    Public Sub InsertAccount(username As String, password As String, email As String)
        Try
            If con.State = ConnectionState.Closed Then con.Open()

            ' Check if Username already exists
            Using checkCmd As New SqlCommand("SELECT COUNT(*) FROM AccountTb WHERE Username = @Username", con)
                checkCmd.Parameters.AddWithValue("@Username", username)
                Dim exists As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                If exists > 0 Then
                    MessageBox.Show("Username already exists. Please choose a different one.", "Duplicate Username", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            End Using

            ' Insert if unique
            Using cmd As New SqlCommand("INSERT INTO AccountTb (Username, Password, Role, Email) VALUES (@Username, @Password, 'Admin', @Email)", con)
                cmd.Parameters.AddWithValue("@Username", username)
                cmd.Parameters.AddWithValue("@Password", password)
                cmd.Parameters.AddWithValue("@Email", email)

                cmd.ExecuteNonQuery()
                MessageBox.Show("Admin account added successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using

        Catch ex As Exception
            MessageBox.Show("Error inserting account: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub



    Public Sub UpdateCandidate(candidateID As Integer, fullName As String, age As String, contactNo As String, email As String,
                           platform As String, yearLevel As String, Section As String, position As String,
                           profilePic As Byte())

        Try
            con.Open()
            Dim query As String = "UPDATE CandidateTb SET 
                                Profile = @Profile,
                                FullName = @FullName,
                                Age = @Age,
                                ContactNo = @ContactNo,
                                Email = @Email,
                                Platform = @Platform,
                                YearLevel = @YearLevel,
                                Section = @Section,
                                Position = @Position
                               WHERE CandidateID = @CandidateID"

            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Profile", If(profilePic, DBNull.Value))
                cmd.Parameters.AddWithValue("@FullName", fullName)
                cmd.Parameters.AddWithValue("@Age", age)
                cmd.Parameters.AddWithValue("@ContactNo", contactNo)
                cmd.Parameters.AddWithValue("@Email", email)
                cmd.Parameters.AddWithValue("@Platform", platform)
                cmd.Parameters.AddWithValue("@YearLevel", yearLevel)
                cmd.Parameters.AddWithValue("@Section", Section)
                cmd.Parameters.AddWithValue("@Position", position)
                cmd.Parameters.AddWithValue("@CandidateID", candidateID)

                cmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Candidate updated successfully!", "Updated", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Public Function Login(username As String, password As String) As Boolean
        Try
            con.Open()

            ' Try to find the matching user
            Dim query As String = "SELECT AccountID, Role FROM AccountTb " &
                              "WHERE Username COLLATE Latin1_General_CS_AS = @Username " &
                              "AND Password COLLATE Latin1_General_CS_AS = @Password"

            Dim accountID As Integer = -1
            Dim role As String = ""

            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Username", username)
                cmd.Parameters.AddWithValue("@Password", password)

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        accountID = Convert.ToInt32(reader("AccountID"))
                        role = reader("Role").ToString()
                    Else
                        reader.Close()

                        ' Failed login: increment FailedAttempts and possibly set LockoutTime
                        Dim updateCmd As New SqlCommand("UPDATE AccountTb SET FailedAttempts = FailedAttempts + 1, LockoutTime = CASE WHEN FailedAttempts + 1 >= 5 THEN @Now ELSE LockoutTime END WHERE Username = @Username", con)
                        updateCmd.Parameters.AddWithValue("@Username", username)
                        updateCmd.Parameters.AddWithValue("@Now", DateTime.Now)
                        updateCmd.ExecuteNonQuery()

                        Return False
                    End If
                End Using
            End Using

            ' Successful login: reset attempts
            Dim resetCmd2 As New SqlCommand("UPDATE AccountTb SET FailedAttempts = 0, LockoutTime = NULL WHERE AccountID = @AccountID", con)
            resetCmd2.Parameters.AddWithValue("@AccountID", accountID)
            resetCmd2.ExecuteNonQuery()

            loggedInRole = role

            ' Fetch voter info
            Dim voterQuery As String = "SELECT VoterID FROM VotersTb WHERE AccountID = @AccountID"
            Using voterCmd As New SqlCommand(voterQuery, con)
                voterCmd.Parameters.AddWithValue("@AccountID", accountID)
                Dim voterResult As Object = voterCmd.ExecuteScalar()
                loggedInUser = If(voterResult IsNot Nothing, voterResult.ToString(), "")
            End Using

            Return True
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try

        Return False
    End Function




    Public Sub FilterCandidate(searchText As String, dgv As DataGridView)
        Try

            con.Open()

            Dim query As String = "SELECT CandidateID, FullName, Age, ContactNo, Email, Platform, YearLevel, Section, Position, Profile, ElectionYear " &
                      "FROM CandidateTb WHERE FullName LIKE @searchText "

            Dim cmd As New SqlCommand(query, con)
            cmd.Parameters.AddWithValue("@searchText", "%" & searchText & "%")

            Dim adapter As New SqlDataAdapter(cmd)
            Dim dt As New DataTable()
            adapter.Fill(dt)
            If dt.Rows.Count > 0 Then
                dgv.DataSource = dt
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message)
        End Try
    End Sub

    Public Sub FilterVoters(searchText As String, dgv As DataGridView)
        Try

            con.Open()

            Dim query As String = "SELECT V.VoterID, A.Username, A.Password, A.Email, V.Fullname, V.YearLevel, V.Section,  " &
                              "CASE WHEN V.HasVoted = 1 THEN 'Yes' ELSE 'No' END AS HasVoted " &
                              "FROM VotersTb V " &
                              "INNER JOIN AccountTb A ON V.AccountID = A.AccountID where V.Fullname like @searchText"


            Dim cmd As New SqlCommand(query, con)
            cmd.Parameters.AddWithValue("@searchText", "%" & searchText & "%")

            Dim adapter As New SqlDataAdapter(cmd)
            Dim dt As New DataTable()
            adapter.Fill(dt)

            ' Replace actual password values with asterisks (same number of characters)
            For Each row As DataRow In dt.Rows
                Dim pwd As String = row("Password").ToString()
                row("Password") = New String("*"c, pwd.Length)
            Next

            If dt.Rows.Count > 0 Then
                dgv.DataSource = dt
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message)
        End Try
    End Sub


End Module
