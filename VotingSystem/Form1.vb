Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form1

    Dim failedAttempts As Integer = 0 ' Counter for failed attempts
    Dim lockoutTime As DateTime? = Nothing ' Store the lockout time when the textboxes are disabled
    Dim lockoutDuration As Integer = 5 ' Lockout duration in minutes (5 minutes)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadEmailToComboBox()
        Panel4.Visible = False
        TextBox1.Text = "Username"
        TextBox1.ForeColor = Color.DarkGray
        TextBox2.Text = "Password"
        TextBox2.ForeColor = Color.DarkGray
        TextBox5.Text = "App Password"
        TextBox5.ForeColor = Color.DarkGray

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim username As String = TextBox1.Text
        Dim password As String = TextBox2.Text

        ' Check if account is locked
        Dim lockoutInfo As (IsLocked As Boolean, RemainingTime As TimeSpan) = CheckLockout(username)
        If lockoutInfo.IsLocked Then
            MessageBox.Show("Account is locked. Try again in " & lockoutInfo.RemainingTime.Minutes & " minutes and " & lockoutInfo.RemainingTime.Seconds & " seconds.", "Locked Out", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Login(username, password) Then
            MessageBox.Show("Login Successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Clear()

            If loggedInRole = "Admin" Then
                Dim adminForm As New Form2()
                adminForm.Show()
                Me.Hide()
            ElseIf loggedInRole = "Voter" Then
                Dim voterForm As New VotingForm()
                voterForm.Show()
                Me.Hide()
            Else
                MessageBox.Show("Invalid Role!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBox.Show("Invalid Username or Password!", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Clear()
        End If
    End Sub

    Private Function CheckLockout(username As String) As (Boolean, TimeSpan)
        Try
            con.Open()
            Dim cmd As New SqlCommand("SELECT FailedAttempts, LockoutTime FROM AccountTb WHERE Username = @Username", con)
            cmd.Parameters.AddWithValue("@Username", username)

            Using reader As SqlDataReader = cmd.ExecuteReader()
                If reader.Read() Then
                    Dim attempts As Integer = Convert.ToInt32(reader("FailedAttempts"))
                    Dim lockoutTime As Object = reader("LockoutTime")

                    If attempts >= 5 AndAlso lockoutTime IsNot DBNull.Value Then
                        Dim lockoutStart As DateTime = Convert.ToDateTime(lockoutTime)
                        Dim elapsed As TimeSpan = DateTime.Now - lockoutStart
                        If elapsed.TotalMinutes < 2 Then
                            Return (True, TimeSpan.FromMinutes(2) - elapsed)
                        Else
                            ' Lockout expired, reset in DB
                            reader.Close()
                            Dim resetCmd As New SqlCommand("UPDATE AccountTb SET FailedAttempts = 0, LockoutTime = NULL WHERE Username = @Username", con)
                            resetCmd.Parameters.AddWithValue("@Username", username)
                            resetCmd.ExecuteNonQuery()
                            Return (False, TimeSpan.Zero)
                        End If
                    End If
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try

        Return (False, TimeSpan.Zero)
    End Function

    ' Timer Tick event to check if the lockout period has passed
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If lockoutTime IsNot Nothing Then
            Dim elapsedTime As TimeSpan = DateTime.Now - lockoutTime.Value
            If elapsedTime.TotalMinutes >= lockoutDuration Then
                TextBox1.Enabled = True
                TextBox2.Enabled = True
                failedAttempts = 0 ' Reset the failed attempts counter
                lockoutTime = Nothing ' Reset lockout time
                Timer1.Stop() ' Stop the timer
                MessageBox.Show("You can now try logging in again.", "Unlock", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button6.PerformClick()
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Panel4.Visible = True
        Panel2.Visible = False
        PictureBox6.Visible = False

        '      PictureBox1.Visible = False

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Panel4.Visible = False
        Panel2.Visible = True
        PictureBox6.Visible = True
        '     PictureBox1.Visible = True
    End Sub


    Public Sub Clear()
        TextBox1.Clear()
        TextBox2.Clear()

        TextBox5.Clear()
        TextBox1.Focus()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Clear()
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Application.ExitThread()
        End If
    End Sub

    Private Sub LoadEmailToComboBox()
        ComboBox1.Items.Clear()
        Dim query As String = "SELECT DISTINCT Email FROM AccountTb WHERE Email IS NOT NULL AND Email <> ''"
        Dim cmd As New SqlCommand(query, con)

        Try
            con.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader()
            While reader.Read()
                ComboBox1.Items.Add(reader("Email").ToString())
            End While
            reader.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading emails: " & ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim email As String = ComboBox1.Text.Trim()
        Dim appPassword As String = TextBox5.Text.Trim()

        If SendExistingPassword(email, appPassword) Then
            ComboBox1.SelectedIndex = -1
            TextBox5.Clear()
        Else
            MessageBox.Show("Please provide valid details", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ComboBox1.SelectedIndex = -1
            TextBox5.Clear()
        End If
    End Sub

    Private Function SendExistingPassword(email As String, appPassword As String) As Boolean
        Dim query As String = "SELECT Password FROM AccountTb WHERE Email = @Email"
        Try
            con.Open()
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Email", email)
                Dim password As Object = cmd.ExecuteScalar()
                If password IsNot Nothing Then
                    SendEmail(email, "Your Password", "Your password is: " & password.ToString(), appPassword)
                    Return True
                Else
                    Return False
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while retrieving the password.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally
            con.Close()
        End Try
    End Function



    Private Sub SendEmail(toEmail As String, subject As String, body As String, appPassword As String)

        Dim smtpClient As New SmtpClient("smtp.gmail.com", 587) With {
            .Credentials = New Net.NetworkCredential(toEmail, appPassword),
            .EnableSsl = True
        }


        Dim mailMessage As New MailMessage() With {
.From = New MailAddress(toEmail),
            .Subject = subject,
            .Body = body,
            .IsBodyHtml = False
        }

        mailMessage.To.Add(toEmail)

        Try
            smtpClient.Send(mailMessage)
            MessageBox.Show("Password sent successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Please provide valid details", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Panel4_Paint(sender As Object, e As PaintEventArgs) Handles Panel4.Paint

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        TextBox2.PasswordChar = If(TextBox2.PasswordChar = "*"c, ControlChars.NullChar, "*"c)
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        TextBox5.PasswordChar = If(TextBox5.PasswordChar = "*"c, ControlChars.NullChar, "*"c)
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click

    End Sub
    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        If TextBox1.Text = "Username" Then
            TextBox1.Text = ""
            TextBox1.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TextBox1_lostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        If TextBox1.Text = "" Then
            TextBox1.Text = "Username"
            TextBox1.ForeColor = Color.DarkGray
        End If
    End Sub
    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        If TextBox2.Text = "Password" Then
            TextBox2.Text = ""
            TextBox2.PasswordChar = "*"
            TextBox2.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TextBox2_lostFocus(sender As Object, e As EventArgs) Handles TextBox2.LostFocus
        If TextBox2.Text = "" Then
            TextBox2.Text = "Password"
            TextBox2.PasswordChar = ""
            TextBox2.ForeColor = Color.DarkGray
        End If
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        If TextBox5.Text = "App Password" Then
            TextBox5.Text = ""
            TextBox5.PasswordChar = "*"
            TextBox5.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TextBox5_lostFocus(sender As Object, e As EventArgs) Handles TextBox5.LostFocus
        If TextBox5.Text = "" Then
            TextBox5.Text = "App Password"
            TextBox5.PasswordChar = ""
            TextBox5.ForeColor = Color.DarkGray
        End If
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class
