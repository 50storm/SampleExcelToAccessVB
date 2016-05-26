Imports System.IO
Imports System.Data
Imports System.Data.OleDb

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Enabled = False


        Dim path As String = "C:\Users\higarashi\Documents\Visual Studio 2015\Projects\SampleExcelToAccess\SampleExcelToAccess\"
        Dim fileName As String = "nihonusd_Apr04_Naka.csv"
        Dim strConn As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=C:\Users\higarashi\Documents\Visual Studio 2015\Projects\SampleExcelToAccess\SampleExcelToAccess\Database1.accdb"


        Dim reader As New StreamReader(path & fileName)
        Dim SqlIns As String = "INSERT INTO Prices (Product,PriceInUSD)VALUES(@Product,@PriceInUSD) "
        Debug.WriteLine("////////////////////////////開始します////////////////////////////")
        Debug.WriteLine(Now)


        Using connection As New OleDbConnection(strConn)
            connection.Open()


            Dim i As ULong = 0
            While (Not reader.EndOfStream)

                Dim line As String = reader.ReadLine()
                i = i + 1
                If (i = 1) Then


                Else
                    '登録
                    Dim para() As String = line.Split(",")
                    Dim Command As New OleDbCommand

                    Try

                        Command.Connection = connection
                        Command.CommandText = SqlIns
                        Command.Parameters.Add("@Product", OleDbType.Variant).Value = para(0).Trim()
                        Command.Parameters.Add("@PriceInUSD", OleDbType.Variant).Value = para(1).Trim()
                        Command.ExecuteNonQuery()

                        'Debug.Write("件数:")
                        'Debug.WriteLine(i.ToString())

                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                        MessageBox.Show(ex.Message)
                    End Try
                    'Debug.WriteLine(line)


                End If
            End While
        End Using


        Debug.WriteLine("//////////////////////終了しました。//////////////////////")
        Debug.WriteLine(Now)

        MessageBox.Show("登録完了")

        Button1.Enabled = True


    End Sub

    'Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    'Dim builder As New OleDbConnectionStringBuilder()
    '    'builder.DataSource = "C:\Users\higarashi\Documents\Visual Studio 2015\Projects\SampleExcelToAccess\SampleExcelToAccess\nihonusd_Apr04_Naka.xlsx"
    '    'builder.Provider = "Microsoft.Jet.Oledb.4.0"

    '    'Console.WriteLine(builder.ConnectionString)

    '    'Dim strConn As String = "Provider =Microsoft.ACE.OLEDB.12.0;" &
    '    '    "Data Source=C:\Users\higarashi\Documents\Visual Studio 2015\Projects\SampleExcelToAccess\SampleExcelToAccess\nihonusd_Apr04_Naka.xlsx;" &
    '    '    "Extended Properties=""Excel 12.0;" & "HDR=YES;" & " IMEX=1;"""

    '    Dim strConn As String = "Provider =Microsoft.ACE.OLEDB.12.0;" &
    '        "Data Source=C:\Users\higarashi\Documents\Visual Studio 2015\Projects\SampleExcelToAccess\SampleExcelToAccess\test.xlsx;" &
    '        "Extended Properties=""Excel 12.0;" & "HDR=YES;" & " IMEX=1;"""


    '    Debug.WriteLine(strConn)

    '    'Using connection As New OleDbConnection(builder.ConnectionString)
    '    '    Try
    '    '        connection.Open()

    '    '        ' Do something with the database here. 
    '    '    Catch ex As Exception
    '    '        Console.WriteLine(ex.Message)
    '    '    End Try
    '    'End Using


    '    Using connection As New OleDbConnection(strConn)
    '        Dim Command As New OleDbCommand
    '        Dim Adapter As New OleDbDataAdapter
    '        Dim dt As New DataTable
    '        Try
    '            connection.Open()
    '            Debug.WriteLine("接続しました")

    '            Command.Connection = connection
    '            'Command.CommandText = "SELECT * FROM [nihonusd1$]"
    '            Command.CommandText = "SELECT * FROM [Sheet1$]"
    '            '規定30秒
    '            '0:無限
    '            Command.CommandTimeout = 0


    '            Adapter.SelectCommand = Command
    '            Adapter.Fill(dt)
    '            gViewExcel.DataSource = dt

    '            ' Do something with the database here. 
    '        Catch ex As Exception
    '            Console.WriteLine(ex.Message)
    '        End Try
    '    End Using

    '    ''http://qiita.com/unarist/items/6cc35bb9fe502ced332f
    '    ''http://blog.sorceryforce.net/?p=154
    'End Sub
End Class
