Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class AddCandidate
    Dim selectedID As Integer = 0
    Private Sub AddCandidate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel1.Visible = False
        LoadCandidate()
        LoadYearLevels()
        LoadSections()
        LoadPositions()
        TextBox6.Text = "Search by Fullname"
        TextBox6.ForeColor = Color.DarkGray

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Panel1.Visible = Not Panel1.Visible
    End Sub

    ' Insert Button
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Collect values from form controls
        Dim fullName As String = TextBox1.Text
        Dim age As String = TextBox2.Text
        Dim contactNo As String = TextBox3.Text
        Dim email As String = TextBox4.Text
        Dim platform As String = TextBox5.Text
        Dim yearLevel As String = ComboBox1.Text
        Dim section As String = ComboBox2.Text
        Dim position As String = ComboBox3.Text
        Dim profilePic As Byte() = ImageToByteArray(PictureBox1.Image)

        Try
            con.Open()

            ' Check for duplicate before inserting
            If IsDuplicate(fullName, yearLevel) Then
                MessageBox.Show("A candidate with the same Full Name already exists in this year level.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            ' SQL Insert (ElectionYear uses YEAR(GETDATE()))
            Dim insertQuery As String = "
            INSERT INTO CandidateTb 
            (Profile, FullName, Age, ContactNo, Email, Platform, YearLevel, Section, Position, VotesCount, ElectionYear) 
            VALUES (@Profile, @FullName, @Age, @ContactNo, @Email, @Platform, @YearLevel, @Section, @Position, 0, YEAR(GETDATE()))"

            Using cmd As New SqlCommand(insertQuery, con)
                cmd.Parameters.AddWithValue("@Profile", If(profilePic, DBNull.Value))
                cmd.Parameters.AddWithValue("@FullName", fullName)
                cmd.Parameters.AddWithValue("@Age", age)
                cmd.Parameters.AddWithValue("@ContactNo", contactNo)
                cmd.Parameters.AddWithValue("@Email", email)
                cmd.Parameters.AddWithValue("@Platform", platform)
                cmd.Parameters.AddWithValue("@YearLevel", yearLevel)
                cmd.Parameters.AddWithValue("@Section", section)
                cmd.Parameters.AddWithValue("@Position", position)

                cmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Candidate inserted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadCandidate()
            ClearData()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            con.Close()
        End Try
    End Sub


    ' Update Button
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim fullName As String = TextBox1.Text
        Dim age As String = TextBox2.Text
        Dim contactNo As String = TextBox3.Text
        Dim email As String = TextBox4.Text
        Dim platform As String = TextBox5.Text
        Dim yearLevel As String = ComboBox1.Text
        Dim section As String = ComboBox2.Text
        Dim position As String = ComboBox3.Text
        Dim profilePic As Byte() = ImageToByteArray(PictureBox1.Image)

        If selectedID <> 0 Then
            Try
                con.Open()

                ' Optional: You can skip the duplicate check if editing same name
                If IsDuplicate(fullName, yearLevel, selectedID) Then
                    MessageBox.Show("Another candidate with the same Full Name exists in this year level.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                ' SQL Update
                Dim query As String = "UPDATE CandidateTb SET 
                                   Profile = @Profile, 
                                   FullName = @FullName, 
                                   Age = @Age, 
                                   ContactNo = @ContactNo, 
                                   Email = @Email, 
                                   Platform = @Platform, 
                                   YearLevel = @YearLevel, 
                                   Section = @Section, 
                                   Position = @Position, 
                                   ElectionYear = YEAR(GETDATE())
                                   WHERE CandidateID = @CandidateID"

                Using cmd As New SqlCommand(query, con)
                    cmd.Parameters.AddWithValue("@Profile", If(profilePic, DBNull.Value))
                    cmd.Parameters.AddWithValue("@FullName", fullName)
                    cmd.Parameters.AddWithValue("@Age", age)
                    cmd.Parameters.AddWithValue("@ContactNo", contactNo)
                    cmd.Parameters.AddWithValue("@Email", email)
                    cmd.Parameters.AddWithValue("@Platform", platform)
                    cmd.Parameters.AddWithValue("@YearLevel", yearLevel)
                    cmd.Parameters.AddWithValue("@Section", section)
                    cmd.Parameters.AddWithValue("@Position", position)
                    cmd.Parameters.AddWithValue("@CandidateID", selectedID)

                    cmd.ExecuteNonQuery()
                End Using

                MessageBox.Show("Candidate updated successfully!", "Updated", MessageBoxButtons.OK, MessageBoxIcon.Information)
                LoadCandidate()
                ClearData()

            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Finally
                con.Close()
            End Try
        Else
            MessageBox.Show("Please select a candidate to update.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    ' Function to check for duplicate entry
    Private Function IsDuplicate(fullName As String, yearLevel As String, Optional excludeID As Integer = 0) As Boolean
        Dim checkQuery As String = "SELECT COUNT(*) FROM CandidateTb WHERE FullName = @FullName AND YearLevel = @YearLevel"
        If excludeID > 0 Then
            checkQuery &= " AND CandidateID <> @CandidateID"
        End If

        Using checkCmd As New SqlCommand(checkQuery, con)
            checkCmd.Parameters.AddWithValue("@FullName", fullName)
            checkCmd.Parameters.AddWithValue("@YearLevel", yearLevel)
            If excludeID > 0 Then
                checkCmd.Parameters.AddWithValue("@CandidateID", excludeID)
            End If

            Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
            Return count > 0
        End Using
    End Function


    Public Function ImageToByteArray(image As Image) As Byte()
        Using ms As New System.IO.MemoryStream()
            ' Clone the image to avoid issues with GDI+ stream locks
            Using cloneImage As New Bitmap(image)
                cloneImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            End Using
            Return ms.ToArray()
        End Using
    End Function



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With New OpenFileDialog
            .Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif"
            If .ShowDialog() = DialogResult.OK Then
                PictureBox1.Image = Image.FromFile(.FileName)
            End If
        End With
    End Sub

    Private Sub LoadCandidate()
        Try
            Dim query As String = "SELECT CandidateID, FullName, Age, ContactNo, Email, Platform, YearLevel, Section, Position, Profile, ElectionYear FROM CandidateTb"
            Dim whereClause As String = ""
            Dim cmd As SqlCommand

            ' Build WHERE clause only if filters are enabled
            If Not CheckBox1.Checked Then
                If ComboBox4.SelectedIndex <> -1 Then whereClause &= " AND YearLevel = @YearLevel"
                If ComboBox5.SelectedIndex <> -1 Then whereClause &= " AND Section = @Section"
                If ComboBox6.SelectedIndex <> -1 Then whereClause &= " AND Position = @Position"
            End If

            If whereClause <> "" Then query &= " WHERE 1=1" & whereClause
            query &= " ORDER BY CandidateID DESC"

            cmd = New SqlCommand(query, con)

            ' Now add parameters if needed
            If Not CheckBox1.Checked Then
                If ComboBox4.SelectedIndex <> -1 Then cmd.Parameters.AddWithValue("@YearLevel", ComboBox4.Text)
                If ComboBox5.SelectedIndex <> -1 Then cmd.Parameters.AddWithValue("@Section", ComboBox5.Text)
                If ComboBox6.SelectedIndex <> -1 Then cmd.Parameters.AddWithValue("@Position", ComboBox6.Text)
            End If

            Using da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)
                DataGridView1.DataSource = dt

                If dt.Rows.Count = 0 Then
                    MessageBox.Show("No candidate records found for the selected filter(s).", "No Records", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

        Catch ex As Exception
            MessageBox.Show("Error loading candidates: " & ex.Message)
        End Try
    End Sub



    Private Sub LoadYearLevels()
        Try
            Dim query As String = "SELECT DISTINCT YearLevel FROM CandidateTb"
            Dim cmd As New SqlCommand(query, con)
            con.Open()

            Dim reader As SqlDataReader = cmd.ExecuteReader()
            ComboBox4.Items.Clear()

            While reader.Read()
                ComboBox4.Items.Add(reader("YearLevel").ToString())
            End While

            reader.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading year levels: " & ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Private Sub LoadSections()
        Try
            Dim query As String = "SELECT DISTINCT Section FROM CandidateTb"
            Dim cmd As New SqlCommand(query, con)
            con.Open()

            Dim reader As SqlDataReader = cmd.ExecuteReader()
            ComboBox5.Items.Clear()

            While reader.Read()
                ComboBox5.Items.Add(reader("Section").ToString())
            End While

            reader.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading sections: " & ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

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

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        LoadCandidate()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        LoadCandidate()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        LoadCandidate()
    End Sub




    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            Panel1.Visible = True


            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)


            TextBox1.Text = row.Cells("FullName").Value.ToString()
            TextBox2.Text = row.Cells("Age").Value.ToString()
            TextBox3.Text = row.Cells("ContactNo").Value.ToString()
            TextBox4.Text = row.Cells("Email").Value.ToString()
            TextBox5.Text = row.Cells("Platform").Value.ToString()
            ComboBox1.Text = row.Cells("YearLevel").Value.ToString()
            ComboBox2.Text = row.Cells("Section").Value.ToString()
            ComboBox3.Text = row.Cells("Position").Value.ToString()


            selectedID = Convert.ToInt32(row.Cells("CandidateID").Value)

            If Not IsDBNull(row.Cells("Profile").Value) Then
                Dim imgBytes As Byte() = CType(row.Cells("Profile").Value, Byte())
                Using ms As New System.IO.MemoryStream(imgBytes)
                    PictureBox1.Image = Image.FromStream(ms)
                End Using
            Else
                PictureBox1.Image = Nothing
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        LoadCandidate()
        CheckBox1.Checked = True
        ComboBox6.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        ComboBox5.SelectedIndex = -1
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ClearData()
        Panel1.Visible = False
        DataGridView1.ClearSelection()
    End Sub
    Public Sub ClearData()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        PictureBox1.Image = Nothing

    End Sub


    Public Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Module1.FilterCandidate(TextBox6.Text, DataGridView1)
    End Sub
    Private Sub TextBox6_GotFocus(sender As Object, e As EventArgs) Handles TextBox6.GotFocus
        If TextBox6.Text = "Search by Fullname" Then
            TextBox6.Text = ""
            TextBox6.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TextBox6_lostFocus(sender As Object, e As EventArgs) Handles TextBox6.LostFocus
        If TextBox6.Text = "" Then
            TextBox6.Text = "Search by Fullname"
            TextBox6.ForeColor = Color.DarkGray
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        LoadCandidate()
    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class