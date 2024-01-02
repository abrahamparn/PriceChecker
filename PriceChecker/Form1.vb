Imports System.IO
Imports IDM.Fungsi
Imports System.Globalization

Imports MySql.Data.MySqlClient

Public Class Form1
    ' Notes. karena simulasi, update row id di matiin. Fungsi write to file di gabungin ke 
    Private stopwatch As New Stopwatch()
    Private whichFunction As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        addToConstCRI()
        ' InitializeProgressBar()
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
    'Private Sub InitializeProgressBar()
    '    ProgressBar1.Maximum = 6000 ' Set this to the same value as the timer interval
    '    ProgressBar1.Step = 1
    '    stopwatch.Start() ' Start the stopwatch
    '    ' Start a timer to update the progress bar
    '    Dim progressTimer As New Timer()
    '    AddHandler progressTimer.Tick, AddressOf UpdateProgressBar
    '    progressTimer.Interval = 1 ' Update progress every millisecond
    '    progressTimer.Start()
    'End Sub

    Private Sub UpdateProgressBar(sender As Object, e As EventArgs)
        ' Update the progress bar based on the stopwatch elapsed time
        If stopwatch.ElapsedMilliseconds <= ProgressBar1.Maximum Then
            ProgressBar1.Value = CInt(stopwatch.ElapsedMilliseconds)
        Else
            ProgressBar1.Value = ProgressBar1.Maximum
            stopwatch.Stop()
            ResetUIAndShowCompletionMessage()

        End If
    End Sub

    Private Sub BGWorker_CheckPrice_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BGWorker_CheckPrice.DoWork
        If whichFunction = 1 Then
            priceChecker()
        ElseIf whichFunction = 2 Then
            RecipeCheckerSub()
        End If
    End Sub

    Private Sub priceChecker()
        Dim sb As New System.Text.StringBuilder
        Dim dt As New DataTable

        sb.AppendLine("Daftar Harga yang Berbeda di STMAST")
        sb.AppendLine()
        sb.AppendLine("PLU - Nama Barang - RowId")
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
                mtran.hpp > mtran.price AND
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
                                sb.AppendLine($"{plu}, {desc2}, {rowId}")
                                Debug.WriteLine(plu)

                                ' Update last processed rowId
                                ' UpdateLastProcessedRowId(rowId)
                            End While
                        End Using
                        WritingLogToFile("PriceChecker", sb.ToString())

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

    Private Sub DoCalculation_Click(sender As Object, e As EventArgs) Handles DoCalculation.Click
        whichFunction = 1 'to call background worker
        stopwatch.Stop()
        stopwatch.Reset()
        ProgressBar1.Value = 0

        ProgressBar1.Maximum = 2500
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
    Private Sub ResetUIAndShowCompletionMessage()

        ProgressBar1.Value = 0
        stopwatch.Stop()
        stopwatch.Reset()

        MessageBox.Show("Background worker has completed its task.", "Task Completed", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub RecipeChecker_Click(sender As Object, e As EventArgs) Handles RecipeChecker.Click
        whichFunction = 2 'to call background worker
        stopwatch.Stop()
        stopwatch.Reset()
        ProgressBar1.Value = 0

        ProgressBar1.Maximum = 2500
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

    Private Sub RecipeCheckerSub()
        Dim da As New MySqlDataAdapter
        Dim dt As New DataTable
        Dim rmplu As String = ""
        Dim total_rmplu As Integer = 0
        Dim sb As New System.Text.StringBuilder
        Try
            Using connection As MySqlConnection = MasterMcon.Clone()
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If

                Using cmd As New MySqlCommand("", connection)
                    Try
                        cmd.CommandText = "Drop table if exists recipe_pricechecker; "
                        cmd.CommandText &= $"
                            CREATE TABLE `recipe_pricechecker` (
                              `RECID` CHAR(1) NOT NULL DEFAULT '',
                              `PLU` VARCHAR(8) NOT NULL DEFAULT '',
                              `RMPLU` VARCHAR(8) NOT NULL DEFAULT '',
                              `DESKRIPSI_RESEP` VARCHAR(50) NOT NULL DEFAULT '',
                              `QTY` DECIMAL(7,2) NOT NULL DEFAULT '0.00',
                              `UNIT` VARCHAR(4) DEFAULT NULL,
                              `ACOST` DECIMAL(15,6) DEFAULT '0.000000',
                              `TOTAL_HPP` DECIMAL(15,6) DEFAULT '0.000000',
                              `BKP` ENUM('N','Y') DEFAULT 'N',
                              `SUB_BKP` CHAR(1) DEFAULT NULL,
                              `FLAGPROD` VARCHAR(2000) DEFAULT NULL,
                              PRIMARY KEY (`PLU`,`RMPLU`)
                            ) ENGINE=INNODB DEFAULT CHARSET=latin1;
                        "
                        cmd.ExecuteNonQuery()

                        cmd.CommandText = $"Select distinct plu from recipe; "
                        da.SelectCommand = cmd
                        da.Fill(dt)

                        sb.AppendLine("Daftar PLU yang Hilang di Prodmast")
                        sb.AppendLine()
                        sb.AppendLine("PLU JUAL - PLU RESEP")
                        For i As Integer = 0 To dt.Rows.Count - 1
                            cmd.CommandText = $"SELECT GROUP_CONCAT(rmplu) FROM recipe WHERE plu = '{dt.Rows(i)("plu")}';"

                            rmplu = "'" & cmd.ExecuteScalar.ToString.Replace(",", "','") & "'"

                            cmd.CommandText = $"SELECT COUNT(rmplu) FROM recipe WHERE plu ='{dt.Rows(i)("plu")}';"
                            total_rmplu = cmd.ExecuteScalar

                            cmd.CommandText = $"SELECT COUNT(prdcd) FROM prodmast WHERE prdcd IN ({rmplu});"
                            If total_rmplu <> cmd.ExecuteScalar Then
                                cmd.CommandText = $"SELECT GROUP_CONCAT(rmplu) FROM recipe WHERE plu = '{dt.Rows(i)("plu")}' AND rmplu NOT IN (SELECT prdcd FROM prodmast);"

                                sb.AppendLine(dt.Rows(i)("plu").ToString() & " - " & "'" & cmd.ExecuteScalar.ToString.Replace(",", "','") & "'")
                            Else
                                cmd.CommandText = $"INSERT INTO recipe_pricechecker 
                                    SELECT r.recid, r.plu, r.rmplu, p.desc2, r.qty, p.unit, p.acost, 
                                    (p.acost * r.qty) total_hpp, p.bkp, p.sub_bkp,
                                    p.flagprod 
                                    FROM recipe r INNER JOIN prodmast p ON r.rmplu = p.prdcd
                                    WHERE p.recid = '' AND p.bkp = 'Y' AND 
                                    p.sub_bkp NOT IN ('A', 'B', 'D', 'L', 'P', 'R', 'S', 'T', 'W', 'G')
                                    AND r.plu = '{dt.Rows(i)("plu").ToString}';
                                    "
                                cmd.ExecuteNonQuery()
                            End If
                        Next

                        WritingLogToFile("RMPLU_HILANG", sb.ToString())

                        cmd.CommandText = $"
                            SELECT t.plu, SUM(t.total_hpp) total_hpp, p.price, p.desc2,
                            (CASE WHEN SUM(t.total_hpp) > p.price THEN 1
                            ELSE 0
                            END ) AS result
                            FROM recipe_pricechecker t INNER JOIN prodmast p ON t.plu = p.prdcd
                            WHERE p.recid = '' 
                            GROUP BY t.plu;
                            "
                        Dim dt_result As New DataTable
                        da.SelectCommand = cmd
                        da.Fill(dt_result)

                        sb = New System.Text.StringBuilder
                        sb.AppendLine("Daftar PLU Resep HPP > Harga Jual")
                        sb.AppendLine()
                        sb.AppendLine("PRDCD - Deskripsi - Total HPP Resep - Harga Jual ")
                        For j As Integer = 0 To dt_result.Rows.Count - 1
                            If dt_result.Rows(j)("total_hpp") > dt_result.Rows(j)("price") Then
                                sb.AppendLine(dt_result.Rows(j)("plu") & " - " &
                                    dt_result.Rows(j)("desc2") & " - " &
                                    dt_result.Rows(j)("total_hpp") & " - " &
                                    dt_result.Rows(j)("price"))
                            End If
                        Next

                        WritingLogToFile("PLU_HPP-VS-PRICE", sb.ToString())


                    Catch ex As Exception
                        TraceLog(ex.Message)
                        MsgBox(ex.Message)
                    End Try
                End Using
                connection.Close()
            End Using
        Catch ex As Exception
            TraceLog(ex.Message)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub WritingLogToFile(title As String, content As String)
        Try
            Dim culture As New CultureInfo("id-ID")
            Dim dateString = DateTime.Now.ToString("ddMMMMyyyy", culture)
            Dim folderPath As String = "D:\LogPCG_Dump"

            ' Check the folder 
            If Not System.IO.Directory.Exists(folderPath) = True Then
                System.IO.Directory.CreateDirectory(folderPath)
            End If

            Dim filePath As String = Path.Combine(folderPath, $"{title}_{dateString}.txt")

            If System.IO.File.Exists(filePath) = True Then
                System.IO.File.Delete(filePath)
            End If

            ' Append the plu to the file
            Using writer As New StreamWriter(filePath, True)
                writer.WriteLine($"{content}")
            End Using
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            TraceLog(ex.Message)
            MsgBox(ex.Message)
        End Try
    End Sub

    ' Fungsi ini di panggil kalo sudah melewati simulasi, gak boleh di hapus file nya karena row idnya di track. 
    Private Sub WritingLogToFilePriceChecker(title As String, content As String)
        Try
            Dim culture As New CultureInfo("id-ID")
            Dim dateString = DateTime.Now.ToString("ddMMMMyyyy", culture)
            Dim folderPath As String = "D:\LogPCG_Dump"

            ' Check the folder 
            If Not System.IO.Directory.Exists(folderPath) = True Then
                System.IO.Directory.CreateDirectory(folderPath)
            End If

            Dim filePath As String = Path.Combine(folderPath, $"{title}_{dateString}.txt")

            'If System.IO.File.Exists(filePath) = True Then
            '    System.IO.File.Delete(filePath)
            'End If

            ' Append the plu to the file
            Using writer As New StreamWriter(filePath, True)
                writer.WriteLine($"{content}")
            End Using
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            TraceLog(ex.Message)
            MsgBox(ex.Message)
        End Try
    End Sub
    'Private Sub BGWorker_CheckPrice_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGWorker_CheckPrice.RunWorkerCompleted
    '    ResetUIAndShowCompletionMessage()
    'End Sub
End Class
