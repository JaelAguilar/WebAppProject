Imports System.Data.SqlClient

Partial Class WebForm1
    Private Function GenerateSQLCommand(sql As SqlCommand, r As DataRow)
        Dim initialQuery As String = String.Empty
        Dim databaseName = DropDownList1.SelectedValue
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

End Class
