Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms.DataVisualization.Charting

Public Class Result

    Private Sub Result_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadPositions()


        ComboBox1.SelectedIndex = -1

    End Sub

    Private Sub LoadPositions()
        Try
            con.Open()
            Dim query As String = "SELECT DISTINCT Position FROM CandidateTb"
            Dim cmd As New SqlCommand(query, con)


            Dim reader As SqlDataReader = cmd.ExecuteReader()
            ComboBox1.Items.Clear()

            While reader.Read()
                ComboBox1.Items.Add(reader("Position").ToString())
            End While

            reader.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading positions: " & ex.Message)
        Finally
            con.Close()
        End Try
    End Sub



    ' Trigger loading on any combo selection
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        LoadFilteredResults()
    End Sub



    Private Sub LoadFilteredResults()
        Try
            Dim query As String = "
            SELECT CandidateID, FullName, VotesCount, Platform, Position 
            FROM CandidateTb 
            WHERE 1=1"

            Dim cmd As New SqlCommand With {
            .Connection = con
        }

            If ComboBox1.SelectedIndex <> -1 Then
                query &= " AND Position = @Position"
                cmd.Parameters.AddWithValue("@Position", ComboBox1.Text)
            End If


            query &= " ORDER BY VotesCount DESC"
            cmd.CommandText = query

            Dim dt As New DataTable()
            Using adapter As New SqlDataAdapter(cmd)
                adapter.Fill(dt)
            End Using

            If dt.Rows.Count = 0 Then
                MessageBox.Show("No records found for the selected filters.", "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
                DataGridView1.DataSource = Nothing
                Chart1.Series.Clear()
                Return
            End If

            ' Bind to DataGridView
            DataGridView1.DataSource = dt

            ' Bind to Chart
            Chart1.Series.Clear()
            Dim series As New Series("Votes") With {
            .ChartType = SeriesChartType.Bar,
            .XValueMember = "FullName",
            .YValueMembers = "VotesCount"
        }
            Chart1.Series.Add(series)

            Chart1.DataSource = dt
            Chart1.DataBind()

            ' Chart styling
            With Chart1.ChartAreas(0).AxisX
                .Interval = 1
                .LabelStyle.Angle = -45
            End With

        Catch ex As Exception
            MessageBox.Show("Error loading results: " & ex.Message)
        Finally
            con.Close()
        End Try
    End Sub


    'Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
    '    If DataGridView1.Rows.Count = 0 Then
    '        MessageBox.Show("No data to export.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '        Return
    '    End If

    '    Dim folderDialog As New FolderBrowserDialog With {
    '        .Description = "Select folder to save the CSV and Chart"
    '    }

    '    If folderDialog.ShowDialog() = DialogResult.OK Then
    '        Dim folderPath = folderDialog.SelectedPath
    '        Dim csvPath = Path.Combine(folderPath, "FilteredResults.csv")
    '        Dim chartPath = Path.Combine(folderPath, "ChartImage.png")

    '        Try
    '            ' === SAVE CSV FILE ===
    '            Using sw As New StreamWriter(csvPath, False, System.Text.Encoding.UTF8)
    '                ' Write headers
    '                Dim headers = DataGridView1.Columns.Cast(Of DataGridViewColumn)().
    '                    Select(Function(col) $"""{col.HeaderText}""")
    '                sw.WriteLine(String.Join(",", headers))

    '                ' Write rows
    '                For Each row As DataGridViewRow In DataGridView1.Rows
    '                    If Not row.IsNewRow Then
    '                        Dim cells = row.Cells.Cast(Of DataGridViewCell)().
    '                            Select(Function(cell) $"""{cell.Value?.ToString().Replace("""", """""")}""")
    '                        sw.WriteLine(String.Join(",", cells))
    '                    End If
    '                Next
    '            End Using

    '            ' === SAVE CHART IMAGE ===
    '            Chart1.SaveImage(chartPath, DataVisualization.Charting.ChartImageFormat.Png)

    '            MessageBox.Show("CSV and Chart image saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

    '        Catch ex As Exception
    '            MessageBox.Show($"Error saving files: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Try
    '    End If
    'End Sub

End Class
