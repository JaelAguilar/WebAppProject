Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading
Imports System.Windows.Forms
Imports ExcelDataReader

Public Class WebForm1
    Inherits Page

    Private Function SaveData(r As DataRow)
        Dim mysqlCOn As New SqlConnection("server=WIN-2VFJL7TQ7Q8\SQLEXPRESS;database=WebDatabase;User ID='externaluser';Password='public12##'")
        Dim mysqlCmd As SqlCommand
        Dim resul As Boolean
        Dim initialQuery As String = String.Empty
        initialQuery &= "INSERT INTO A_1_PresupuestoGlobal (Secretaria,Dirección,I_PreEst,I_PreMod,I_PreDev,I_PreRec,EPreOrigApro,E_1A_AmpPres,E_2A_AmpPres,E_3A_AmpPres,E_Tot_Amp,E_PreModif,E_PreComp,E_PreDev,E_PreEjer,E_PreErog,E_PreCons,E_PrePorEjer,FechaCorte,Elaboró,Revisó,Autorizó)"

        initialQuery &= "VALUES (@sec,@dir,@IpreEst,@IpreMod,@IpreDev,@IpreRec,@EorigApro,@1ampPres,@2ampPres,@3ampPres,@EtotAmp,@EpreModif,@EpreComp,@EpreDev,@EpreEjer,@EpreErog,@EpreCons,@EprePorEjer,@FCorte,@elab,@rev,@aut)"
        Try
            mysqlCOn.Open()
            mysqlCmd = New SqlCommand
            With mysqlCmd
                .Connection = mysqlCOn
                .CommandType = CommandType.Text
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
                resul = .ExecuteNonQuery()
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            mysqlCOn.Close()
        End Try
        Return resul
    End Function

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
End Class