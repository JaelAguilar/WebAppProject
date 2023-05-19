Imports System.Data.SqlClient
Imports System.Threading
Imports System.IO
Imports System.Windows.Forms
Imports ExcelDataReader




Public Class WebForm1
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub

    Protected Sub importExcel_Click(sender As Object, e As EventArgs) Handles importExcel.Click
        Dim thread As New Thread(
            Sub()
                Dim fileDial As New OpenFileDialog With {
            .Filter = "Excel Files|*.xls; *.xlsx; *.xlsm"
        }
                If fileDial.ShowDialog() = DialogResult.Cancel Then
                    Debug.Write("Failed retrieving the file")
                    Return
                End If
                Debug.Write("Retrieved file")
                Dim stream2 As New FileStream(fileDial.FileName, FileMode.Open)
                Dim xslReader As IExcelDataReader
                xslReader = ExcelReaderFactory.CreateOpenXmlReader(stream2)
                Dim table = xslReader.AsDataSet().Tables(0)
                Dim tableRows = table.Rows
                tableRows.RemoveAt(0)
                For Each r As DataRow In tableRows
                    Debug.Write("Secretaría" & r(0).ToString)
                    Dim resul = SaveData(r)
                Next
                Debug.WriteLine("SQL Table updated")
                stream2.Close()
                xslReader.Close()
            End Sub
            )
        thread.SetApartmentState(ApartmentState.STA)
        thread.Start()
    End Sub

    Private Function SaveData(r As DataRow)
        Dim mysqlCOn As New SqlConnection("server=WIN-2VFJL7TQ7Q8\SQLEXPRESS;database=WebDatabase;User ID='externaluser';Password='public12##'")
        Dim mysqlCmd As SqlCommand
        Dim resul As Boolean
        Try
            mysqlCOn.Open()
            mysqlCmd = New SqlCommand
            With mysqlCmd
                .Connection = mysqlCOn
                .CommandType = CommandType.Text
            End With
            GenerateSQLCommand(mysqlCmd, r)
            resul = mysqlCmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            mysqlCOn.Close()
        End Try
        Return resul
    End Function

End Class