Imports System.Data.SqlClient
Imports System.Threading
Imports System.IO
Imports System.Windows.Forms
Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports ExcelDataReader
Imports MigraDoc
Imports MigraDoc.DocumentObjectModel
Imports MigraDoc.Rendering
Imports MigraDoc.DocumentObjectModel.Tables
Imports Microsoft

Public Class WebForm1
    Inherits Page

    Private r As New List(Of Integer)({1000, 2000, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16})

    <Obsolete>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim pdfDoc = CreatePDF()
        Dim renderer As New PdfDocumentRenderer(True, PdfFontEmbedding.Always) With {
.Document = pdfDoc
}
        renderer.RenderDocument()
        Dim loc = "C:\Users\admin\Downloads\test.pdf"
        renderer.PdfDocument.Save(loc)
        Process.Start(loc)
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

    Private Function GenerateSQLCommand(sql As SqlCommand, r As DataRow)
        Dim initialQuery As String = String.Empty
        Dim databaseName = importTableSelector.SelectedValue
        Debug.WriteLine("DATABASE = " + databaseName)

        Select Case databaseName
            Case "A.1"
                initialQuery &= "INSERT INTO A.1 (Secretaria,Dirección,I_PreEst,I_PreMod,I_PreDev,I_PreRec,EPreOrigApro,E_1A_AmpPres,E_2A_AmpPres,E_3A_AmpPres,E_Tot_Amp,E_PreModif,E_PreComp,E_PreDev,E_PreEjer,E_PreErog,E_PreCons,E_PrePorEjer,FechaCorte,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@IpreEst,@IpreMod,@IpreDev,@IpreRec,@EorigApro,@1ampPres,@2ampPres,@3ampPres,@EtotAmp,@EpreModif,@EpreComp,@EpreDev,@EpreEjer,@EpreErog,@EpreCons,@EprePorEjer,@FCorte,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@IpreEst", r(2))
                    .Parameters.AddWithValue("@IpreMod", r(3))
                    .Parameters.AddWithValue("@IpreDev", r(4))
                    .Parameters.AddWithValue("@IpreRec", r(5))
                    .Parameters.AddWithValue("@EorigApro", r(6))
                    .Parameters.AddWithValue("@1ampPres", r(7))
                    .Parameters.AddWithValue("@2ampPres", r(8))
                    .Parameters.AddWithValue("@3ampPres", r(9))
                    .Parameters.AddWithValue("@EtotAmp", r(10))
                    .Parameters.AddWithValue("@EpreModif", r(11))
                    .Parameters.AddWithValue("@EpreComp", r(12))
                    .Parameters.AddWithValue("@EpreDev", r(13))
                    .Parameters.AddWithValue("@EpreEjer", r(14))
                    .Parameters.AddWithValue("@EpreErog", r(15))
                    .Parameters.AddWithValue("@EpreCons", r(16))
                    .Parameters.AddWithValue("@EprePorEjer", r(17))
                    .Parameters.AddWithValue("@FCorte", r(18))
                    .Parameters.AddWithValue("@elab", r(19))
                    .Parameters.AddWithValue("@rev", r(20))
                    .Parameters.AddWithValue("@aut", r(21))
                End With
        End Select
    End Function
    Private Function ExportPDF()

    End Function

    Protected Sub generateReport_Click(sender As Object, e As EventArgs) Handles generateReport.Click
        Dim pdfDoc = CreatePDF()
        Dim renderer As New PdfDocumentRenderer(True) With {
            .Document = pdfDoc
        }
        renderer.RenderDocument()

        ' Saving the pdf
        Dim saveFileDialog As New SaveFileDialog With {
            .Filter = "PDF document (*.pdf)|*.pdf",
            .Title = "Guardar el reporte"
        }
        Dim thread As New Thread(
            Sub()
                If saveFileDialog.ShowDialog() = DialogResult.OK Then
                    Dim pdfFilename As String = saveFileDialog.FileName
                    renderer.Save(pdfFilename)
                    Process.Start(pdfFilename)
                Else
                    MessageBox.Show("Hubo un error, inténtelo de nuevo")
                End If
            End Sub
            )
        thread.SetApartmentState(ApartmentState.STA)
        thread.Start()

    End Sub



End Class