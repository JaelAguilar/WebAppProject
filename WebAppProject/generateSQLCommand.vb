Imports System.Data.SqlClient
Imports System.Security.Cryptography

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

            Case "A.1.1"
                initialQuery &= "INSERT INTO A.1.1
(Secretaria, Dirección,Impuestos,Cuot_ApSS,Cont_Mej,Derechos,Productos,Aprov,Ing_Vta_Bs,Part_Apor,Tras_Sub_Oayu,Ing_Finan,ClaveEjercicio,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Direccion,@Impuestos,@Cuot_ApSS,@Cont_Mej,@Derechos,@Productos,@Aprov,@Ing_Vta_Bs,@Part_Apor,@Tras_Sub_Oayu,@Ing_Finan,@ClaveEjercicio,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Direccion", r(0))
                    .Parameters.AddWithValue("@Impuestos", r(1))
                    .Parameters.AddWithValue("@Cuot_ApSS", r(2))
                    .Parameters.AddWithValue("@Cont_Mej", r(3))
                    .Parameters.AddWithValue("@Derechos", r(4))
                    .Parameters.AddWithValue("@Producto", r(5))
                    .Parameters.AddWithValue("@Aprov", r(6))
                    .Parameters.AddWithValue("@Ing_Vta_Bs", r(7))
                    .Parameters.AddWithValue("@Part_Apor", r(8))
                    .Parameters.AddWithValue("@Tras_Sub_Oayu", r(9))
                    .Parameters.AddWithValue("@Ing_Finan", r(10))
                    .Parameters.AddWithValue("@ClaveEjercicio", r(11))
                    .Parameters.AddWithValue("@Elaboró", r(12))
                    .Parameters.AddWithValue("@Revisó", r(13))
                    .Parameters.AddWithValue("@Autorizó", r(14))
                End With

            Case "A.1.2"
                initialQuery &= "INSERT INTO A.1.2
(Secretaria, Dirección,Serv_Per,Mat_Sum,Serv_Gen,Tras_Sub_Oayu,Bien_Mu_Inm_Inta,Iver_Pub,Inver_Fin_OP,Part_Apor,Dedua_Pub,FechaCorte,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Secretaria,@Dirección,@Serv_Per,@Mat_Sum,@Serv_Gen,@Tras_Sub_Oayu,@Bien_Mu_Inm_Inta,@Iver_Pub,Inver_Fin_OP,@Part_Apor,@Dedua_Pub,@FechaCorte,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@Serv_Per", r(2))
                    .Parameters.AddWithValue("@Mat_Sum", r(3))
                    .Parameters.AddWithValue("@Serv_Gen", r(4))
                    .Parameters.AddWithValue("@Tras_Sub_Oayu", r(5))
                    .Parameters.AddWithValue("@Bien_Mu_Inm_Inta", r(6))
                    .Parameters.AddWithValue("@Iver_Pub,Inver_Fin_OP", r(7))
                    .Parameters.AddWithValue("@Part_Apor", r(8))
                    .Parameters.AddWithValue("@Dedua_Pub", r(9))
                    .Parameters.AddWithValue("@FechaCorte", r(10))
                    .Parameters.AddWithValue("@Elaboró", r(11))
                    .Parameters.AddWithValue("@Revisó", r(12))
                    .Parameters.AddWithValue("@Autorizó", r(13))
                End With

            Case "A.1.3"
                initialQuery &= "INSERT INTO A.1.3
(Secretaria,Dirección,Clave,Nombre,Presup_Auto,Porcentaje,FechaCorte,Elaboró,Revisó)"
                initialQuery &= "VALUES
(@Secretaria,@Dirección,@Clave,@Nombre,@Presup_Auto,@Porcentaje,@FechaCorte,@Elaboró,@Revisó,)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@ecretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@Clave", r(2))
                    .Parameters.AddWithValue("@Nombre", r(3))
                    .Parameters.AddWithValue("@Presup_Auto", r(4))
                    .Parameters.AddWithValue("@Porcentaje", r(5))
                    .Parameters.AddWithValue("@FechaCorte", r(6))
                    .Parameters.AddWithValue("@Elaboró", r(7))
                    .Parameters.AddWithValue("@Revisó", r(8))
                End With
        End Select
    End Function

End Class
