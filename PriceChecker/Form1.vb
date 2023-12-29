Imports System.IO
Imports IDM.Fungsi
Imports MySql.Data.MySqlClient

Public Class Form1
    Private stopwatch As New Stopwatch()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        addToConstCRI()
        ' InitializeProgressBar() '
    End Sub
    'Private Sub Timer_CheckPrice_Tick(sender As Object, e As EventArgs) Handles Timer_CheckPrice.Tick
    '    ' Reset progress bar when the timer ticks
    '    ProgressBar1.Value = 0

    '    stopwatch.Stop()
    '    stopwatch.Reset()
    '    stopwatch.Start() ' Restart the stopwatch for a new cycle
    '    If Not BGWorker_CheckPrice.IsBusy Then
    '        BGWorker_CheckPrice.RunWorkerAsync()
    '    End If
    'End Sub
    Private Sub InitializeProgressBar()
        ProgressBar1.Maximum = 6000 ' Set this to the same value as the timer interval
        ProgressBar1.Step = 1
        stopwatch.Start() ' Start the stopwatch
        ' Start a timer to update the progress bar
        Dim progressTimer As New Timer()
        AddHandler progressTimer.Tick, AddressOf UpdateProgressBar
        progressTimer.Interval = 1 ' Update progress every millisecond
        progressTimer.Start()
    End Sub

    Private Sub UpdateProgressBar(sender As Object, e As EventArgs)
        ' Update the progress bar based on the stopwatch elapsed time
        If stopwatch.ElapsedMilliseconds <= ProgressBar1.Maximum Then
            ProgressBar1.Value = CInt(stopwatch.ElapsedMilliseconds)
        Else
            ProgressBar1.Value = ProgressBar1.Maximum
            stopwatch.Stop()
        End If
    End Sub

    Private Sub BGWorker_CheckPrice_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BGWorker_CheckPrice.DoWork
        Using connection As MySqlConnection = MasterMcon.Clone()
            If connection.State = ConnectionState.Closed Then
                connection.Open()
            End If

            ' Get the starting rowId
            Dim startRowId As Integer = Convert.ToDecimal(GetStartingRowId())

            ' Query to get records from mtran where rowId > startRowId
            Dim query As String = $"
           
            SELECT 
                prodmast.desc2, 
                mtran.plu, 
                mtran.rowid
            FROM mtran 
            INNER JOIN prodmast ON prodmast.prdcd = mtran.plu 
            WHERE 
                mtran.rtype = 'J' AND mtran.rowid > {startRowId} AND  
                ((mtran.gross_dpp+mtran.PPN ) / mtran.qty) > mtran.price AND
                 prodmast.BKP = 'Y' AND 
                prodmast.SUB_BKP NOT IN ('A', 'B', 'D', 'L', 'P', 'R', 'S', 'T', 'W', 'G')
            GROUP BY mtran.docno, mtran.plu, mtran.shift, mtran.station, mtran.tanggal
                 ORDER BY mtran.rowId ASC
            
            "

            Try
                Using command As New MySqlCommand(query, connection)
                    Try
                        Using reader As MySqlDataReader = command.ExecuteReader()
                            While reader.Read()
                                Dim plu As String = reader("plu").ToString()
                                Dim desc2 As String = reader("desc2").ToString()
                                Dim rowId As Decimal = Convert.ToDecimal(reader("rowid"))
                                WriteToFile(plu, desc2, rowId)
                                Debug.WriteLine(plu)

                                ' Update last processed rowId
                                UpdateLastProcessedRowId(rowId)
                            End While
                        End Using
                    Catch ex As Exception
                        Debug.WriteLine(ex.Message)
                        TraceLog(ex.Message)
                        MsgBox(ex.Message)

                    End Try
                End Using
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
                TraceLog(ex.Message)
                MsgBox(ex.Message)

            End Try
        End Using
    End Sub
    Private Sub addToConstCRI()
        Try
            Using connection As MySqlConnection = MasterMcon.Clone()
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If

                Using command As New MySqlCommand()
                    command.Connection = connection

                    command.CommandText = "SELECT COUNT(*) FROM const WHERE rkey = 'CRI'"
                    Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())

                    If count = 0 Then
                        command.CommandText = "INSERT INTO const (RKEY, `DESC`, DOCNO) VALUES ('CRI', '0', 0)"
                        command.ExecuteNonQuery()
                    End If
                End Using
            End Using
        Catch ex As Exception
            TraceLog(ex.Message)
            MsgBox(ex.Message)

        End Try

    End Sub
    Private Function GetStartingRowId() As Integer
        Dim startRowId As Integer = 0
        Try
            Using connection As MySqlConnection = MasterMcon.Clone()
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                Dim query As String = $"SELECT const.desc from const where Rkey = 'CRI'"

                Using command As New MySqlCommand(query, connection)
                    Try
                        Using reader As MySqlDataReader = command.ExecuteReader()
                            While reader.Read()
                                startRowId = Convert.ToInt32(reader("desc"))
                            End While
                        End Using
                    Catch ex As Exception
                        Debug.WriteLine(ex.Message)
                        TraceLog(ex.Message)
                        MsgBox(ex.Message)

                    End Try
                End Using
            End Using
        Catch ex As Exception
            TraceLog(ex.Message)
            MsgBox(ex.Message)
        End Try
        Return startRowId
    End Function

    Private Sub UpdateLastProcessedRowId(rowId As Integer)
        Try
            Using connection As MySqlConnection = MasterMcon.Clone()
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If

                Dim query As String = $"UPDATE  const SET const.DESC = '{rowId}' where Rkey = 'CRI'"

                Using command As New MySqlCommand(query, connection)
                    Try
                        command.ExecuteScalar()
                    Catch ex As Exception
                        TraceLog(ex.Message)
                        MsgBox(ex.Message)

                    End Try
                End Using
            End Using
        Catch ex As Exception
            TraceLog(ex.Message)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub WriteToFile(plu As String, desc2 As String, rowId As Integer)
        Try
            Dim documentsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim filePath As String = Path.Combine(documentsPath, "CheckPriceResult.txt")

            ' Append the plu to the file
            Using writer As New StreamWriter(filePath, True) ' True to append data
                writer.WriteLine($"{plu}, {desc2}, {rowId}")
            End Using
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            TraceLog(ex.Message)
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub DoCalculation_Click(sender As Object, e As EventArgs) Handles DoCalculation.Click
        stopwatch.Stop()
        stopwatch.Reset()
        ProgressBar1.Value = 0

        ProgressBar1.Maximum = 6000
        ProgressBar1.Step = 1
        stopwatch.Start()
        Dim progressTimer As New Timer()
        AddHandler progressTimer.Tick, AddressOf UpdateProgressBar

        progressTimer.Interval = 1 ' Update progress every millisecond
        progressTimer.Start()
        If Not BGWorker_CheckPrice.IsBusy Then
            BGWorker_CheckPrice.RunWorkerAsync()
        End If
    End Sub
End Class
