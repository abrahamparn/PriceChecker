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
        CheckMasterResepTempTable()
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

    Private Sub doWorkNow()
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

    Private Sub BGWorker_CheckPrice_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BGWorker_CheckPrice.DoWork
        If whichFunction = 1 Then
            priceChecker()
        ElseIf whichFunction = 2 Then
            RecipeCheckerSub()
        ElseIf whichFunction = 3 Then
            If CSVReader() = False Then
                Exit Sub
            End If
            CSVCheckPluBahanBaku()
            checkPluAsalDanAcost()
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
        doWorkNow()
    End Sub
    Private Sub ResetUIAndShowCompletionMessage()

        ProgressBar1.Value = 0
        stopwatch.Stop()
        stopwatch.Reset()

        MessageBox.Show("Background worker has completed its task.", "Task Completed", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub RecipeChecker_Click(sender As Object, e As EventArgs) Handles RecipeChecker.Click
        whichFunction = 2 'to call background worker
        doWorkNow()

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
    Private Sub CheckMasterResepTempTable()
        Try
            Using connection As MySqlConnection = MasterMcon.Clone()
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If

                Using command As New MySqlCommand()
                    command.Connection = connection
                    command.CommandText = "Drop table if exists resepMasterTemp; "
                    command.CommandText &= $"
                            CREATE TABLE `resepMasterTemp` (
                              `PLU_JUAL` VARCHAR(8) NOT NULL DEFAULT '',
                              `PLU_BAHAN_BAKU` VARCHAR(8) NOT NULL DEFAULT '',
                              `QTY` VARCHAR(8) NOT NULL DEFAULT '',
                              `SATUAN` VARCHAR(50) NOT NULL DEFAULT '',
                              PRIMARY KEY (`PLU_JUAL`,`PLU_BAHAN_BAKU`)
                            ) ENGINE=INNODB DEFAULT CHARSET=latin1;
                        "
                    command.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            TraceLog(ex.Message)
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Function CSVReader()
        Dim folderPath As String = "D:\MasterResep"
        ' Check the folder 
        If Not System.IO.Directory.Exists(folderPath) = True Then
            System.IO.Directory.CreateDirectory(folderPath)
        End If

        Dim files As String() = System.IO.Directory.GetFiles(folderPath)

        Select Case files.Length
            Case 0
                MessageBox.Show("Folder D:\MasterResep tidak boleh kosong. ")
                Return False
            Case > 1
                MessageBox.Show("Folder D:\MasterResep memiliki lebih dari satu file")
                Return False
            Case 1
                If Not files(0).EndsWith(".csv") Then
                    MessageBox.Show("Folder D:\MasterResep hanya boleh menganduk file csv")
                    Return False
                End If

                If files(0).Contains("'") Then
                    MessageBox.Show($"Nama file tidak boleh menganduk petik tunggal")
                    Return False
                End If

                Try
                    Using reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(files(0))
                        reader.TextFieldType = FileIO.FieldType.Delimited
                        reader.SetDelimiters(",")

                        If Not reader.EndOfData Then
                            Dim currentRow As String() = reader.ReadFields()
                            Dim columnCount As Integer = 0

                            For Each field As String In currentRow
                                If Not String.IsNullOrEmpty(field.Trim()) Then ' remove white spaces
                                    columnCount += 1
                                End If
                            Next

                            If columnCount > 4 Then
                                MessageBox.Show("File csv memiliki lebih dari empat column")
                                Return False
                            End If
                        End If
                    End Using


                    Using connection As MySqlConnection = MasterMcon.Clone()
                        If connection.State = ConnectionState.Closed Then
                            connection.Open()
                        End If

                        Using cmd As New MySqlCommand()
                            cmd.Connection = connection
                            cmd.CommandText = "DELETE FROM resepMasterTemp"
                            cmd.ExecuteNonQuery()

                            Dim csvFilePath As String = files(0).Replace("\", "\\")
                            cmd.CommandText = "LOAD DATA LOCAL INFILE '" & csvFilePath &
                                          "' INTO TABLE resepMasterTemp " &
                                          "FIELDS TERMINATED BY ',' " &
                                          "LINES TERMINATED BY '\r\n' " &
                                          "IGNORE 1 LINES " &
                                          "(PLU_JUAL, PLU_BAHAN_BAKU, QTY, SATUAN)"
                            cmd.ExecuteNonQuery()
                        End Using
                        connection.Close()
                    End Using

                    'IO.File.Delete(files(0)) enable this later
                    Return True
                Catch ex As Exception
                    MessageBox.Show("Error reading the CSV file: " & ex.Message)
                    TraceLog("Error reading the CSV file: " & ex.Message)
                    Return False

                End Try
        End Select

    End Function
    Private Sub CSV_Checker_Click(sender As Object, e As EventArgs) Handles CSV_Checker.Click

        whichFunction = 3
        doWorkNow()

    End Sub

    Private Sub CSVCheckPluBahanBaku()
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
                        cmd.CommandText = "DELETE FROM resepMasterTemp WHERE plu_bahan_baku = '-' OR plu_bahan_baku = '' OR plu_bahan_baku = ' ';"
                        cmd.ExecuteScalar()

                        cmd.CommandText = $"Select distinct PLU_JUAL from resepMasterTemp; "
                        da.SelectCommand = cmd
                        da.Fill(dt)

                        sb.AppendLine("Daftar PLU yang Hilang di Prodmast berdasarkan CSV")
                        sb.AppendLine()
                        sb.AppendLine("PLU JUAL - PLU RESEP")
                        For i As Integer = 0 To dt.Rows.Count - 1
                            cmd.CommandText = $"SELECT GROUP_CONCAT(PLU_BAHAN_BAKU) FROM resepMasterTemp WHERE PLU_JUAL = '{dt.Rows(i)("PLU_JUAL")}';"

                            rmplu = "'" & cmd.ExecuteScalar.ToString.Replace(",", "','") & "'"

                            cmd.CommandText = $"SELECT COUNT(PLU_BAHAN_BAKU) FROM resepMasterTemp WHERE PLU_JUAL ='{dt.Rows(i)("PLU_JUAL")}';"
                            total_rmplu = cmd.ExecuteScalar

                            cmd.CommandText = $"SELECT COUNT(prdcd) FROM prodmast WHERE prdcd IN ({rmplu});"
                            If total_rmplu <> cmd.ExecuteScalar Then
                                cmd.CommandText = $"SELECT GROUP_CONCAT(PLU_BAHAN_BAKU) FROM resepMasterTemp " +
                                                  $"WHERE PLU_JUAL = '{dt.Rows(i)("PLU_JUAL")}' " +
                                                  $"And PLU_BAHAN_BAKU Not IN (SELECT prdcd FROM prodmast) ;"

                                sb.AppendLine(dt.Rows(i)("PLU_JUAL").ToString() & " - " & "'" & cmd.ExecuteScalar.ToString.Replace(",", "','") & "'")

                            End If
                        Next
                        sb.AppendLine()
                        sb.AppendLine()
                        sb.AppendLine("Daftar PLU yang Hilang di konversi_plu berdasarkan CSV")
                        sb.AppendLine()
                        sb.AppendLine("PLU JUAL - PLU RESEP")
                        For i As Integer = 0 To dt.Rows.Count - 1
                            cmd.CommandText = $"SELECT GROUP_CONCAT(PLU_BAHAN_BAKU) FROM resepMasterTemp WHERE PLU_JUAL = '{dt.Rows(i)("PLU_JUAL")}';"

                            rmplu = "'" & cmd.ExecuteScalar.ToString.Replace(",", "','") & "'"

                            cmd.CommandText = $"SELECT COUNT(PLU_BAHAN_BAKU) FROM resepMasterTemp WHERE PLU_JUAL ='{dt.Rows(i)("PLU_JUAL")}';"
                            total_rmplu = cmd.ExecuteScalar

                            cmd.CommandText = $"SELECT COUNT(plu_konv) FROM konversi_plu WHERE plu_konv IN ({rmplu});"
                            If total_rmplu > cmd.ExecuteScalar Then
                                cmd.CommandText = $"SELECT GROUP_CONCAT(PLU_BAHAN_BAKU) FROM resepMasterTemp " +
                                                  $"WHERE PLU_JUAL = '{dt.Rows(i)("PLU_JUAL")}' " +
                                                  $"And PLU_BAHAN_BAKU Not IN (SELECT plu_konv FROM konversi_plu) ;"


                                sb.AppendLine(dt.Rows(i)("PLU_JUAL").ToString() & " - " & "'" & cmd.ExecuteScalar.ToString.Replace(",", "','") & "'")

                            End If
                        Next

                        WritingLogToFile("PLU_BAHAN_BAKU_HILANG", sb.ToString())
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

    Private Sub checkPluAsalDanAcost()
        Dim da As New MySqlDataAdapter
        Dim dt As New DataTable
        Dim rmplu As String = ""
        Dim total_rmplu As Integer = 0
                Dim ab As New System.Text.StringBuilder

        Dim sb As New System.Text.StringBuilder
        Try
            Using connection As MySqlConnection = MasterMcon.Clone()
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If

                Using cmd As New MySqlCommand("", connection)
                    Try
                        cmd.CommandText = "DELETE FROM resepMasterTemp WHERE plu_bahan_baku = '-' OR plu_bahan_baku = '' OR plu_bahan_baku = ' ';"
                        cmd.ExecuteScalar()
                        cmd.CommandText = $"SELECT 
                                                DISTINCT r.PLU_BAHAN_BAKU, 
                                                b.plu_konv, 
                                                b.plu_asal, 
                                                p.prdcd AS prdcd_asal, 
                                                p.acost/b.nilai AS hasilBagi_asal, 
                                                z.plu_konv_acost
                                            FROM resepMasterTemp AS r
                                            INNER JOIN konversi_plu AS b ON r.PLU_BAHAN_BAKU = b.plu_konv 
                                            INNER JOIN prodmast AS p ON b.plu_asal = p.prdcd
                                            INNER JOIN (
                                                SELECT 
                                                    r_inner.PLU_BAHAN_BAKU, 
                                                    acost AS plu_konv_acost
                                                FROM resepMasterTemp AS r_inner
                                                INNER JOIN konversi_plu AS b_inner ON r_inner.PLU_BAHAN_BAKU = b_inner.plu_konv 
                                                INNER JOIN prodmast AS p_inner ON b_inner.plu_konv = p_inner.prdcd
                                                GROUP BY r_inner.PLU_BAHAN_BAKU
                                            ) AS z ON z.PLU_BAHAN_BAKU = r.PLU_BAHAN_BAKU"

                        da.SelectCommand = cmd
                        da.Fill(dt)
                        sb.AppendLine("Inserting where acost is not the same as acost_awal/nilai_awal")
                        sb.AppendLine()
                        sb.AppendLine("PLU_BAHAN_BAKU|plu_konv|plu_asal|prdcd_asal|hasilBagi_asal|plu_konv_acost")
                        ab.AppendLine("Inserting query to recipe where acost is the same as acost_awal/nilai_awal")
                        ab.AppendLine()
                        For i As Integer = 0 To dt.Rows.Count - 1
                            If Math.Round(Convert.ToDecimal(dt.Rows(i)("hasilBagi_asal")), 3) <> Math.Round(Convert.ToDecimal(dt.Rows(i)("plu_konv_acost")), 3) Then


                                sb.AppendLine($"{dt.Rows(i)("PLU_BAHAN_BAKU").ToString()}|{dt.Rows(i)("plu_konv").ToString()}|{dt.Rows(i)("plu_asal").ToString()}|{dt.Rows(i)("prdcd_asal").ToString()}|{dt.Rows(i)("hasilBagi_asal").ToString()}|{dt.Rows(i)("plu_konv_acost").ToString()}")
                            Else
                                cmd.CommandText = $"
                                SELECT CONCAT('INSERT INTO recipe (plu, rmplu, qty, addtime) VALUES (\'', plu_jual, '\', \'', plu_bahan_baku, '\', \'', qty, '\', now());') 
                                FROM resepMasterTemp 
                                WHERE plu_bahan_baku = '{dt.Rows(i)("PLU_BAHAN_BAKU")}';                                
                                "
                                ab.AppendLine($"{cmd.ExecuteScalar.ToString}")

                            End If
                        Next

                        WritingLogToFile("Banding_Konv_asal_beda", sb.ToString())
                        WritingLogToFile("insert_to_recipe_query", ab.ToString())

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
    'Private Sub BGWorker_CheckPrice_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGWorker_CheckPrice.RunWorkerCompleted
    '    ResetUIAndShowCompletionMessage()
    'End Sub
End Class
