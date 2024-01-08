﻿Imports MySql.Data.MySqlClient
Imports IDM.Fungsi

Module InitializationProject
    Public MasterMcon As MySqlConnection
    Public isSector As Boolean = False
    Friend a As New IDM.Sector
    Friend MyKey As String = "6D9A2B6195BFD0EB5B7C3ACBEB58AB99"

    Public Sub Main()
        Try
            If isSector Then
                MasterMcon = a.GetVersionV2(MyKey, Application.StartupPath & "\PriceChecker.exe", "kasir")
            Else
                MasterMcon = New MySqlConnection("server=localhost;user id=root;Password=v4dg4IDbVLYJnB7zOv3lKg8jw8WPvrwd4=NqpoGGrLCX;port=3306;database=pos;Allow User Variables=True;")
            End If

            If IsNothing(MasterMcon) Then
                MsgBox("Key Sector Salah!", MsgBoxStyle.Exclamation)
                End
            End If


            Dim f As New Form1
            f.ShowDialog()
            'End If
        Catch ex As Exception
            TraceLog(ex.Message)
            MsgBox(ex.Message)
        End Try

    End Sub
End Module
