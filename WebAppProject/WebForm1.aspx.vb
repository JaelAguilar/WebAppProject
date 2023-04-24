Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Threading

Public Class WebForm1
    Inherits System.Web.UI.Page

    Private Function SaveData(sql As String)
        Dim mysqlCOn As New SqlConnection("server=WIN-2VFJL7TQ7Q8\SQLEXPRESS;database=WebDatabase;User ID='externaluser';Password='public12##'")
        Dim mysqlCmd As SqlCommand
        Dim resul As Boolean

        Try

            mysqlCOn.Open()
            mysqlCmd = New SqlCommand
            With mysqlCmd
                .Connection = mysqlCOn
                .CommandText = sql
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
        Dim OLEcon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & "C:\Users\admin\Downloads\A.1 Presupuesto Global.xlsx" & " ; " & "Extended Properties='Excel 8.0;HDR=Yes'")
        Dim OLEcmd As New OleDbCommand
        Dim OLEda As New OleDbDataAdapter
        Dim OLEdt As New DataTable
        Dim sql As String
        Dim resul As Boolean

        Try

            OLEcon.Open()
            With OLEcmd
                .Connection = OLEcon
                .CommandText = "select * from [A.1$]"
            End With
            Debug.WriteLine("Good")
            OLEda.SelectCommand = OLEcmd
            OLEda.Fill(OLEdt)

            For Each r As DataRow In OLEdt.Rows

                sql = "INSERT INTO A_1_PresupuestoGlobal (Secretaria,Dirección,I_PreEst,I_PreMod,I_PreDev,I_PreRec,EPreOrigApro,E_1A_AmpPres,E_2A_AmpPres,E_3A_AmpPres,E_Tot_Amp,E_PreModif,E_PreComp,E_PreDev,E_PreEjer,E_PreErog,E_PreCons,E_PrePorEjer,FechaCorte,Elaboró,Revisó,Autorizó) VALUES ('" & r(0).ToString & "','" & r(1).ToString & "','" & r(2).ToString & "','" & r(3).ToString & "','" & r(4).ToString & "','" & r(5).ToString & "','" & r(6).ToString & "','" & r(7).ToString & "','" & r(8).ToString & "','" & r(9).ToString & "','" & r(10).ToString & "','" & r(11).ToString & "','" & r(12).ToString & "','" & r(13).ToString & "','" & r(14).ToString & "','" & r(15).ToString & "','" & r(16).ToString & "','" & r(17).ToString & "','" & r(18).ToString & "','" & r(19).ToString & "','" & r(20).ToString & "','" & r(21).ToString & "')"
                resul = SaveData(sql)

            Next
        Catch ex As Exception
            Debug.WriteLine("Failed")
        Finally
            OLEcon.Close()
        End Try









    End Sub
End Class