Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class AddVoter
    Private selectedAccountID As Integer = -1

    Private Sub AddVoter_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadYearLevels()
        LoadSections()
        LoadHasVoted()
        LoadVoters()
        Panel1.Visible = False
        TextBox6.Text = "Search by Fullname"
        TextBox6.ForeColor = Color.DarkGray
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            con.Open()

            If Duplicate() Then
                con.Close()
                Return
            End If

            ' Insert into AccountTb
            Dim insertAccountQuery As String = "INSERT INTO AccountTb (Username, Password, Role, Email) VALUES (@Username, @Password, 'Voter', @Email); SELECT SCOPE_IDENTITY();"
            Using cmd As New SqlCommand(insertAccountQuery, con)
                cmd.Parameters.AddWithValue("@Username", TextBox1.Text)
                cmd.Parameters.AddWithValue("@Password", TextBox2.Text)
                cmd.Parameters.AddWithValue("@Email", TextBox4.Text)

                Dim newAccountID As Integer = Convert.ToInt32(cmd.ExecuteScalar())

                ' Insert into VotersTb
                Dim insertVoterQuery As String = "INSERT INTO VotersTb (AccountID, Fullname, Section, YearLevel, HasVoted) VALUES (@AccountID, @Fullname, @Section, @YearLevel, 0);"
                Using cmd2 As New SqlCommand(insertVoterQuery, con)
                    cmd2.Parameters.AddWithValue("@AccountID", newAccountID)
                    cmd2.Parameters.AddWithValue("@Fullname", TextBox3.Text)
                    cmd2.Parameters.AddWithValue("@Section", ComboBox2.Text)
                    cmd2.Parameters.AddWithValue("@YearLevel", ComboBox1.Text)

                    cmd2.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show("Voter Registered Successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadVoters()

            ' Clear fields
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Registration Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Public Function Duplicate() As Boolean
        ' Check if Username already exists
        Dim checkUsernameQuery As String = "SELECT COUNT(*) FROM AccountTb WHERE Username = @Username"
        Using checkCmd As New SqlCommand(checkUsernameQuery, con)
            checkCmd.Parameters.AddWithValue("@Username", TextBox1.Text)
            Dim usernameExists As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
            If usernameExists > 0 Then
                MessageBox.Show("Username already exists. Please choose another one.", "Duplicate Username", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return True
            End If
        End Using

        ' Check if Fullname already exists in the same YearLevel
        Dim checkFullnameQuery As String = "SELECT COUNT(*) FROM VotersTb WHERE Fullname = @Fullname AND YearLevel = @YearLevel"
        Using checkCmd2 As New SqlCommand(checkFullnameQuery, con)
            checkCmd2.Parameters.AddWithValue("@Fullname", TextBox3.Text)
            checkCmd2.Parameters.AddWithValue("@YearLevel", ComboBox1.Text)
            Dim fullnameExists As Integer = Convert.ToInt32(checkCmd2.ExecuteScalar())
            If fullnameExists > 0 Then
                MessageBox.Show("Fullname already exists in this year level. Please enter a unique name.", "Duplicate Fullname", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return True
            End If
        End Using

        Return False
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            con.Open()

            Dim accountId As Integer = Convert.ToInt32(DataGridView1.CurrentRow.Cells("AccountID").Value)

            ' Check for duplicate Username
            Dim checkUsernameQuery As String = "SELECT COUNT(*) FROM AccountTb WHERE Username = @Username AND AccountID <> @AccountID"
            Using cmdCheckUsername As New SqlCommand(checkUsernameQuery, con)
                cmdCheckUsername.Parameters.AddWithValue("@Username", TextBox1.Text)
                cmdCheckUsername.Parameters.AddWithValue("@AccountID", accountId)

                If Convert.ToInt32(cmdCheckUsername.ExecuteScalar()) > 0 Then
                    MessageBox.Show("Username already exists. Please choose a different one.", "Duplicate Username", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            End Using

            ' Check for duplicate Fullname
            Dim checkFullnameQuery As String = "SELECT COUNT(*) FROM VotersTb WHERE Fullname = @Fullname AND AccountID <> @AccountID"
            Using cmdCheckFullname As New SqlCommand(checkFullnameQuery, con)
                cmdCheckFullname.Parameters.AddWithValue("@Fullname", TextBox3.Text)
                cmdCheckFullname.Parameters.AddWithValue("@AccountID", accountId)

                If Convert.ToInt32(cmdCheckFullname.ExecuteScalar()) > 0 Then
                    MessageBox.Show("Fullname already exists. Please enter a unique name.", "Duplicate Fullname", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            End Using

            ' Update AccountTb
            Dim updateAccountQuery As String = "UPDATE AccountTb SET Username = @Username, Password = @Password, Email = @Email WHERE AccountID = @AccountID"
            Using cmd As New SqlCommand(updateAccountQuery, con)
                cmd.Parameters.AddWithValue("@Username", TextBox1.Text)
                cmd.Parameters.AddWithValue("@Password", TextBox2.Text)
                cmd.Parameters.AddWithValue("@Email", TextBox4.Text)
                cmd.Parameters.AddWithValue("@AccountID", accountId)

                cmd.ExecuteNonQuery()
            End Using

            ' Update VotersTb
            Dim updateVoterQuery As String = "UPDATE VotersTb SET Fullname = @Fullname, Section = @Section, YearLevel = @YearLevel WHERE AccountID = @AccountID"
            Using cmd2 As New SqlCommand(updateVoterQuery, con)
                cmd2.Parameters.AddWithValue("@Fullname", TextBox3.Text)
                cmd2.Parameters.AddWithValue("@Section", ComboBox2.Text)
                cmd2.Parameters.AddWithValue("@YearLevel", ComboBox1.Text)
                cmd2.Parameters.AddWithValue("@AccountID", accountId)

                cmd2.ExecuteNonQuery()
            End Using

            MessageBox.Show("Voter Information Updated Successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadVoters()

            ' Clear fields
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub






    Private Sub LoadVoters()
        Try
            Dim query As String = "Select A.AccountID, V.VoterID, A.Username, A.Password, A.Email, V.Fullname, V.YearLevel, V.Section, " &
                              "CASE WHEN V.HasVoted = 1 THEN 'Yes' ELSE 'No' END AS HasVoted " &
                              "FROM VotersTb V " &
                              "INNER JOIN AccountTb A ON V.AccountID = A.AccountID"
            Dim whereClause As String = ""
            Dim cmd As SqlCommand

            ' Only apply filters if CheckBox1 is not checked
            If Not CheckBox1.Checked Then
                If ComboBox4.SelectedIndex <> -1 Then whereClause &= " AND V.YearLevel = @YearLevel"
                If ComboBox5.SelectedIndex <> -1 Then whereClause &= " AND V.Section = @Section"
                If ComboBox6.SelectedIndex <> -1 Then
                    If ComboBox6.Text = "Yes" Then
                        whereClause &= " AND V.HasVoted = 1"
                    ElseIf ComboBox6.Text = "No" Then
                        whereClause &= " AND V.HasVoted = 0"
                    End If
                End If
            End If

            If whereClause <> "" Then query &= " WHERE 1=1" & whereClause
            query &= " ORDER BY V.VoterID DESC"

            cmd = New SqlCommand(query, con)

            ' Add parameters based on selected filters
            If Not CheckBox1.Checked Then
                If ComboBox4.SelectedIndex <> -1 Then cmd.Parameters.AddWithValue("@YearLevel", ComboBox4.Text)
                If ComboBox5.SelectedIndex <> -1 Then cmd.Parameters.AddWithValue("@Section", ComboBox5.Text)
            End If

            Using da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                ' Mask password
                For Each row As DataRow In dt.Rows
                    Dim pwd As String = row("Password").ToString()
                    row("Password") = New String("*"c, pwd.Length)
                Next

                DataGridView1.DataSource = dt
                DataGridView1.Columns("AccountID").Visible = False
                DataGridView1.Columns("VoterID").Visible = False


                If dt.Rows.Count = 0 Then
                    MessageBox.Show("No voter records found for the selected filter(s).", "No Records", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

        Catch ex As Exception
            MessageBox.Show("Error loading voters: " & ex.Message)
        End Try
    End Sub



    Private Sub LoadYearLevels()
        Try
            Dim query As String = "SELECT DISTINCT YearLevel FROM VotersTb"
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
            Dim query As String = "SELECT DISTINCT Section FROM VotersTb"
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

    Private Sub LoadHasVoted()
        Try
            ComboBox6.Items.Clear()
            ComboBox6.Items.Add("Yes")
            ComboBox6.Items.Add("No")
        Catch ex As Exception
            MessageBox.Show("Error loading vote status: " & ex.Message)
        End Try
    End Sub



    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        LoadVoters()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        LoadVoters()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        LoadVoters()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        LoadVoters()
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            ' Make sure the clicked row index is valid
            If e.RowIndex >= 0 Then
                Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

                Panel1.Visible = True

                ' Populate the fields from the selected row
                TextBox1.Text = row.Cells("Username").Value.ToString()
                TextBox2.Text = row.Cells("Password").Value.ToString()
                TextBox3.Text = row.Cells("Fullname").Value.ToString()
                TextBox4.Text = row.Cells("Email").Value.ToString()
                ComboBox2.Text = row.Cells("Section").Value.ToString()
                ComboBox1.Text = row.Cells("YearLevel").Value.ToString()
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading data from selected row: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Module1.FilterVoters(TextBox6.Text, DataGridView1)
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
    Private Sub ClearForm()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ClearForm()
        Panel1.Visible = False
        DataGridView1.ClearSelection()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Panel1.Visible = Not Panel1.Visible
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        LoadVoters()
        CheckBox1.Checked = True
        ComboBox6.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        ComboBox5.SelectedIndex = -1
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub
End Class