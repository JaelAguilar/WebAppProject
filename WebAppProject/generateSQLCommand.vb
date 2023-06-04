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

            Case "A.1.4"
                initialQuery &= "INSERT INTO A.1.4
(Secretaria,Dirección,Presup_Auto,1a_AmpPre,2a_AmpPre,3a_AmpPre,Total_Amp,Pre_Modificado,Pre_Comprometido,Pre_Devengado,Pre_Ejercicio,Pre_Erogado,Pre_Consuido,Pre_PorEjercer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Secretaria,@Dirección,@Presup_Auto,@1a_AmpPre,@2a_AmpPre,@3a_AmpPre,@Total_Amp,@Pre_Modificado,@Pre_Comprometido,@Pre_Devengado,@Pre_Ejercicio,@Pre_Erogado,@Pre_Consuido,@Pre_PorEjercer,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaría", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@Presup_Auto", r(2))
                    .Parameters.AddWithValue("@1a_AmpPre", r(3))
                    .Parameters.AddWithValue("@2a_AmpPre", r(4))
                    .Parameters.AddWithValue("@3a_AmpPre", r(5))
                    .Parameters.AddWithValue("@Total_Amp", r(6))
                    .Parameters.AddWithValue("@Pre_Modificado", r(7))
                    .Parameters.AddWithValue("@Pre_Modificado", r(8))
                    .Parameters.AddWithValue("@Pre_Comprometido", r(9))
                    .Parameters.AddWithValue("@Pre_Devengado", r(10))
                    .Parameters.AddWithValue("@Pre_Ejercicio", r(11))
                    .Parameters.AddWithValue("@Pre_Erogado", r(12))
                    .Parameters.AddWithValue("@Pre_Consuido", r(13))
                    .Parameters.AddWithValue("@Pre_PorEjercer", r(14))
                    .Parameters.AddWithValue("@Elaboró", r(15))
                    .Parameters.AddWithValue("@Revisó", r(16))
                    .Parameters.AddWithValue("@Autorizó", r(17))
                End With

            Case "A.2"
                initialQuery &= "INSERT INTO A.2
(Secretaria,Dirección,Presup_Auto,AmpRedu,PreModif,PreComp,Total_Amp,Pre_Devengado,Pre_Ejercicio,Pre_Erogado,Pre_Consumido,Pre_PorEjercer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Sec,@Dir,@Presup_Auto,@AmpRedu,@PreModif,@PreComp,@Total_Amp,@Pre_Devengado,@Pre_Ejercicio,@Pre_Erogado,@Pre_Consumido,@Pre_PorEjercer,@Elab,@Rev,@Aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Sec", r(0))
                    .Parameters.AddWithValue("@Dir", r(1))
                    .Parameters.AddWithValue("@Presup_Auto", r(2))
                    .Parameters.AddWithValue("@AmpRedu", r(3))
                    .Parameters.AddWithValue("@PreModif", r(4))
                    .Parameters.AddWithValue("@PreComp", r(5))
                    .Parameters.AddWithValue("@Total_Amp", r(6))
                    .Parameters.AddWithValue("@Pre_Devengado", r(7))
                    .Parameters.AddWithValue("@Pre_Ejercicio", r(8))
                    .Parameters.AddWithValue("@Pre_Erogado", r(9))
                    .Parameters.AddWithValue("@Pre_Consumido", r(10))
                    .Parameters.AddWithValue("@Pre_PorEjercer", r(11))
                    .Parameters.AddWithValue("@Elab", r(12))
                    .Parameters.AddWithValue("@Rev", r(13))
                    .Parameters.AddWithValue("@Aut", r(14))
                End With

            Case "A.3"
                initialQuery &= "INSERT INTO A.3
(Secretaria,Dirección,ESFDA_BalGen,ESFDA_EdoRes,Cta_Pub,Dic_Auditor,Ult_Per_Dic,Bal_Comp,FechaCorte,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Sec,@Dir,@ESFDABalGen,@ESFDAEdoRes,@CtaPub,@DicAuditor,@UltPerDic,@BalComp,@FechaCorte,@Elab,@Rev,@Aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Sec", r(0))
                    .Parameters.AddWithValue("@Dir", r(1))
                    .Parameters.AddWithValue("@ESFDABalGen", r(2))
                    .Parameters.AddWithValue("@ESFDAEdoRes", r(3))
                    .Parameters.AddWithValue("@CtaPub", r(4))
                    .Parameters.AddWithValue("@DicAuditor", r(5))
                    .Parameters.AddWithValue("@UltPerDic", r(6))
                    .Parameters.AddWithValue("@BalComp", r(7))
                    .Parameters.AddWithValue("@FechaCorte", r(8))
                    .Parameters.AddWithValue("@Elab", r(9))
                    .Parameters.AddWithValue("@Rev", r(10))
                    .Parameters.AddWithValue("@Aut", r(11))
                End With

            Case "A.4"
                initialQuery &= "INSERT INTO A.4
(Secretaria,Dirección,Nom_Secretaria,Titular,Monto,FechaCorte,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Sec,@Dir,@NomSec,@Titular,@Monto,@FechaCorte,@Elab,@Rev,@Aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Sec", r(0))
                    .Parameters.AddWithValue("@Dir", r(1))
                    .Parameters.AddWithValue("@NomSec", r(2))
                    .Parameters.AddWithValue("@Titular", r(3))
                    .Parameters.AddWithValue("@Monte", r(4))
                    .Parameters.AddWithValue("@FechaCorte", r(5))
                    .Parameters.AddWithValue("@Elab", r(6))
                    .Parameters.AddWithValue("@Rev", r(7))
                    .Parameters.AddWithValue("@Aut", r(8))
                End With

            Case "A.4.1a"
                initialQuery &= "INSERT INTO A.4.1a
(Secretaria,Dirección,Fecha,Resp_Fondo,Encarg_Fondo,Monto_Aut,Total_Efvo,Total_Docs,Total_Arq,Variación,Billetes_1000,Billetes_500,Billetes_200,Billetes_100,Billetes_50,Monedas_20,Monedas_10,Monedas_5,Monedas_2,Monedas_1,Monedas_50c,Monedas_20c,Monedas_10c,Observaciones,FechaCorte,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Sec,@Dir,@Fecha,@RespFondo,@EncargFondo,@MontoAut,@TotalEfvo,@TotalDocs,@TotalArq,@Var,@Bill1000,@Bill500,@Bill200,@Bill100,@Bill50,@Mon20,@Mon10,@Mon5,@Mon2,@Mon1,@Mon50c,@Mon20c,@Mon10c,@Obs,@FechaCorte,@Elab,@Rev,@Aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Sec", r(0))
                    .Parameters.AddWithValue("@Dir", r(1))
                    .Parameters.AddWithValue("@Fecha", r(2))
                    .Parameters.AddWithValue("@RespFondo", r(3))
                    .Parameters.AddWithValue("@EncargFondo", r(4))
                    .Parameters.AddWithValue("@MontoAut", r(5))
                    .Parameters.AddWithValue("@TotalEfvo", r(6))
                    .Parameters.AddWithValue("@TotalDocs", r(7))
                    .Parameters.AddWithValue("@TotalArq", r(8))
                    .Parameters.AddWithValue("@Var", r(9))
                    .Parameters.AddWithValue("@Bill1000", r(10))
                    .Parameters.AddWithValue("@Bill500", r(11))
                    .Parameters.AddWithValue("@Bill200", r(12))
                    .Parameters.AddWithValue("@Bill100", r(13))
                    .Parameters.AddWithValue("@Bill50", r(14))
                    .Parameters.AddWithValue("@Mon20", r(15))
                    .Parameters.AddWithValue("@Mon10", r(16))
                    .Parameters.AddWithValue("@Mon5", r(17))
                    .Parameters.AddWithValue("@Mon2", r(18))
                    .Parameters.AddWithValue("@Mon1", r(19))
                    .Parameters.AddWithValue("@Mon50c", r(20))
                    .Parameters.AddWithValue("@Mon20c", r(21))
                    .Parameters.AddWithValue("@Mon10c", r(22))
                    .Parameters.AddWithValue("@Obs", r(23))
                    .Parameters.AddWithValue("@FechaCorte", r(24))
                    .Parameters.AddWithValue("@Elab", r(25))
                    .Parameters.AddWithValue("@Rev", r(26))
                    .Parameters.AddWithValue("@Aut", r(27))
                End With

            Case "A.4.1b"
                initialQuery &= "INSERT INTO A.4.1b
(Secretaria,Direccioón,Fecha,Documento,Fecha_Doc,Proveedor,Concepto,Importe,FechaCorte,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Sec,@Dir,@Fec,@Doc,@FechaDoc,@Pro,@Con,@Imp,@FechaCorte,@Elab,@Rev,@Aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Sec", r(0))
                    .Parameters.AddWithValue("@Dir", r(1))
                    .Parameters.AddWithValue("@Fec", r(2))
                    .Parameters.AddWithValue("@Doc", r(3))
                    .Parameters.AddWithValue("@FechaDoc", r(4))
                    .Parameters.AddWithValue("@Pro", r(5))
                    .Parameters.AddWithValue("@Con", r(6))
                    .Parameters.AddWithValue("@Imp", r(7))
                    .Parameters.AddWithValue("@FechaCorte", r(8))
                    .Parameters.AddWithValue("@Elab", r(9))
                    .Parameters.AddWithValue("@Rev", r(10))
                    .Parameters.AddWithValue("@Aut", r(11))
                End With

            Case "A.5"
                initialQuery &= "INSERT INTO A.5
(Secretaria,Dirección,Nombre_Inst,No_Cta,Tipo,Saldo,Fecha_Venc,FechaCorte,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Sec,@Dir,@NombreInst,@NoCta,@Tip,@Sal,@FechaVenc,@FechaCorte,@Elab,@Rev,@Aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Sec", r(0))
                    .Parameters.AddWithValue("@Dir", r(1))
                    .Parameters.AddWithValue("@NombreInst", r(2))
                    .Parameters.AddWithValue("@NoCta", r(3))
                    .Parameters.AddWithValue("@Tip", r(4))
                    .Parameters.AddWithValue("@Sal", r(5))
                    .Parameters.AddWithValue("@FechaVenc", r(6))
                    .Parameters.AddWithValue("@FechaCorte", r(7))
                    .Parameters.AddWithValue("@Elab", r(8))
                    .Parameters.AddWithValue("@Rev", r(9))
                    .Parameters.AddWithValue("@Aut", r(10))
                End With

            Case "A.5.1"
                initialQuery &= "INSERT INTO A.4.1a
(Secretaria,Dirección,Nom_Inst,Num_Cuenta,Cta_Contable,Saldo_SL,Saldo_SECB,Cheq_Bco_I,Cheq_Bco_F,Firma_R1,Firma_R2,Firma_R3,Firma_R4,Cta_Individual,Cta_Mancomunada,Cta_Indistinta,ClaveEjercicio,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Sec,@Dir,@NomInst,@NumCuenta,@CtaContable,@SaldoSL,@SaldoSECB,@CheqBcoI,@CheqBcoF,@FirmaR1,@FirmaR2,@FirmaR3,@FirmaR4,@CtaIndividual,@CtaMancomunada,@CtaIndistinta,@Cla,@Elab,@Rev,@Aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Sec", r(0))
                    .Parameters.AddWithValue("@Dir", r(1))
                    .Parameters.AddWithValue("@NomInst", r(2))
                    .Parameters.AddWithValue("@NumCuenta", r(3))
                    .Parameters.AddWithValue("@CtaContable", r(4))
                    .Parameters.AddWithValue("@SaldosSL", r(5))
                    .Parameters.AddWithValue("@SaldosSECB", r(6))
                    .Parameters.AddWithValue("@CheqBcoI", r(7))
                    .Parameters.AddWithValue("@CheqBcoF", r(8))
                    .Parameters.AddWithValue("@FirmaR1", r(9))
                    .Parameters.AddWithValue("@FirmaR2", r(10))
                    .Parameters.AddWithValue("@FirmaR3", r(11))
                    .Parameters.AddWithValue("@FirmaR4", r(12))
                    .Parameters.AddWithValue("@CtaIndividual", r(13))
                    .Parameters.AddWithValue("@CtaMancomunada", r(14))
                    .Parameters.AddWithValue("@CtaIndistinta", r(15))
                    .Parameters.AddWithValue("@Cla", r(16))
                    .Parameters.AddWithValue("@Elab", r(17))
                    .Parameters.AddWithValue("@Rev", r(18))
                    .Parameters.AddWithValue("@Aut", r(19))
                End With

            Case "A.6"
                initialQuery &= "INSERT INTO A.6 (Secretaria,Dirección,Num_Documento,Nom_Deudor,Fech_Adeudo,Importe_Tot,Saldo,Vencimiento,Concepto,ClaveEjercicio, Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@NumDocumento,@NomDeudor,@FechAdeudo,@ImporteTot,@Saldo,@Vencimiento,@Concepto, @ClaveEjercicio,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@NumDocumento", r(2))
                    .Parameters.AddWithValue("@NomDeudor", r(3))
                    .Parameters.AddWithValue("@FechAdeudo ", r(4))
                    .Parameters.AddWithValue("@ImporteTot", r(5))
                    .Parameters.AddWithValue("@Saldo", r(6))
                    .Parameters.AddWithValue("@Vencimiento", r(7))
                    .Parameters.AddWithValue("@Concepto", r(8))
                    .Parameters.AddWithValue("@ClaveEjercicio", r(9))
                    .Parameters.AddWithValue("@elab", r(10))
                    .Parameters.AddWithValue("@rev", r(11))
                    .Parameters.AddWithValue("@aut", r(12))
                End With

            Case "A.6.2"
                initialQuery &= "INSERT INTO A.6.2 (Secretaria,Dirección,RSS,ClaveEjercicio, Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@rss,@ClaveEjercicio,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@rss", r(2))
                    .Parameters.AddWithValue("@ClaveEjercicio", r(3))
                    .Parameters.AddWithValue("@elab", r(4))
                    .Parameters.AddWithValue("@rev", r(5))
                    .Parameters.AddWithValue("@aut", r(6))
                End With

            Case "A.7"
                initialQuery &= "INSERT INTO A.7 (Secretaria,Dirección,Num_Documento,Nom_Acreador,Saldo,ClaveEjercicio, Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@NumDocumento,@NomAcreador,@Saldo,@ClaveEjercicio,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@NumDocumento", r(2))
                    .Parameters.AddWithValue("@NomAcreador", r(3))
                    .Parameters.AddWithValue("@Saldo", r(4))
                    .Parameters.AddWithValue("@ClaveEjercicio", r(5))
                    .Parameters.AddWithValue("@elab", r(6))
                    .Parameters.AddWithValue("@rev", r(7))
                    .Parameters.AddWithValue("@aut", r(8))
                End With

            Case "A.7.1"
                initialQuery &= "INSERT INTO A7.1 (Secretaria,Dirección,NomProveedor,Saldo,ClaveEjercicio,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@nomproveedor,@saldo,@claveejercicio,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@nomproveedor", r(2))
                    .Parameters.AddWithValue("@saldo", r(3))
                    .Parameters.AddWithValue("@claveejercicio", r(4))
                    .Parameters.AddWithValue("@elab", r(5))
                    .Parameters.AddWithValue("@rev", r(6))
                    .Parameters.AddWithValue("@aut", r(7))
                End With

            Case "A.7.2"
                initialQuery &= "INSERT INTO A7.2 (Secretaria,Dirección,NomProveedor,Saldo,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@nomproveedor,@saldo,@cvecorteejer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@nomproveedor", r(2))
                    .Parameters.AddWithValue("@saldo", r(3))
                    .Parameters.AddWithValue("@cvecorteejer", r(4))
                    .Parameters.AddWithValue("@elab", r(5))
                    .Parameters.AddWithValue("@rev", r(6))
                    .Parameters.AddWithValue("@aut", r(7))
                End With

            Case "A.7.3"
                initialQuery &= "INSERT INTO A7.3 (Secretaria,Dirección,NoDocumento,NomAcredor,Fecha,Saldo,Vencimiento,Concepto,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@nodocumento,@nomacredor,@fecha,@saldo,@vencimiento,@concepto,@cvecorteejer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@nodocumento", r(2))
                    .Parameters.AddWithValue("@nomacredor", r(3))
                    .Parameters.AddWithValue("@fecha", r(4))
                    .Parameters.AddWithValue("@saldo", r(5))
                    .Parameters.AddWithValue("@vencimiento", r(6))
                    .Parameters.AddWithValue("@concepto", r(7))
                    .Parameters.AddWithValue("@cvecorteejer", r(8))
                    .Parameters.AddWithValue("@elab", r(9))
                    .Parameters.AddWithValue("@rev", r(10))
                    .Parameters.AddWithValue("@aut", r(11))
                End With

            Case "A.8"
                initialQuery &= "INSERT INTO A.8 (Secretaria,Dirección,Fecha,NoCtaBanco,Institución,NoCheque,NomBenif,Importe,CveCorteEjer,Elaboró,Revisó,Autorizo)"
                initialQuery &= "VALUES (@sec,@dir,@Fecha,@NoCtaBan,@Institu,@Nocheque,@Nombenif,@Importe,@CveCorteEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@Fecha", r(2))
                    .Parameters.AddWithValue("@NoCtaBan", r(3))
                    .Parameters.AddWithValue("@Institu", r(4))
                    .Parameters.AddWithValue("@Nocheque", r(5))
                    .Parameters.AddWithValue("@Nombenif", r(6))
                    .Parameters.AddWithValue("@Importe", r(7))
                    .Parameters.AddWithValue("@CveCorteEjer", r(8))
                    .Parameters.AddWithValue("@elab", r(9))
                    .Parameters.AddWithValue("@rev", r(10))
                    .Parameters.AddWithValue("@aut", r(11))
                End With

            Case "A.9"
                initialQuery &= "INSERT INTO A.9 (Secretaria,Dirección,NoPoliza,NomAfianzadora,NomDeudor,Monto,ConceptoFianza,CveCorteEjer,Elaboró,Revisó,Autorizo)"
                initialQuery &= "VALUES (@sec,@dir,@NoPoliza,@NomAfianzadora,@NomDeudor,@Monto,@ConceptoFianza,@CveCorteEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@NoPoliza", r(2))
                    .Parameters.AddWithValue("@NomAfianzadora", r(3))
                    .Parameters.AddWithValue("@NomDeudor", r(4))
                    .Parameters.AddWithValue("@Monto", r(5))
                    .Parameters.AddWithValue("@ConceptoFianza", r(6))
                    .Parameters.AddWithValue("@CveCorteEjer", r(7))
                    .Parameters.AddWithValue("@elab", r(8))
                    .Parameters.AddWithValue("@rev", r(9))
                    .Parameters.AddWithValue("@aut", r(10))
                End With

            Case "A.10"
                initialQuery &= "INSERT INTO A.10 (Secretaria,Dirección,No,EjercicioFiscal,ActaEAESNL,NoLegajos,NoDiscos,Estatus,Observaciones,Responsables,CveCorteEjer,Elaboró,Revisó,Autorizo)"
                initialQuery &= "VALUES (@sec,@dir,@No,@EjercicioFiscal,@ActaEAESNL,@NoLegajos,@NoDiscos,@Estatus,@Observaciones,@Responsables,@CveCorteEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@No", r(2))
                    .Parameters.AddWithValue("@EjercicioFiscal", r(3))
                    .Parameters.AddWithValue("@ActaEAESNL", r(4))
                    .Parameters.AddWithValue("@NoLegajos", r(5))
                    .Parameters.AddWithValue("@NoDiscos", r(6))
                    .Parameters.AddWithValue("@Estatus ", r(7))
                    .Parameters.AddWithValue("@Observaciones", r(8))
                    .Parameters.AddWithValue("@Responsables", r(9))
                    .Parameters.AddWithValue("@CveCorteEjer", r(10))
                    .Parameters.AddWithValue("@elab ", r(11))
                    .Parameters.AddWithValue("@rev ", r(12))
                    .Parameters.AddWithValue("@aut ", r(13))
                End With

            Case "B.1"
                initialQuery &= "INSERT INTO B.1 (Secretaria,Dirección,No_Nomina,Puesto,NombreCompleto,Sueldo,Vigencia,Sindicalizado,RegimenEmpleado,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@No_Nomina,@Puesto,@NombreCompleto,@Sueldo,@Vigencia,@Sindicalizado,@RegimenEmpleado,@CveCorteEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@No_Nomina", r(2))
                    .Parameters.AddWithValue("@Puesto", r(3))
                    .Parameters.AddWithValue("@NombreCompleto", r(4))
                    .Parameters.AddWithValue("@Sueldo", r(5))
                    .Parameters.AddWithValue("@Vigencia", r(6))
                    .Parameters.AddWithValue("@Sindicalizado", r(7))
                    .Parameters.AddWithValue("@RegimenEmpleado", r(8))
                    .Parameters.AddWithValue("@CveCorteEjer", r(9))
                    .Parameters.AddWithValue("@elab", r(10))
                    .Parameters.AddWithValue("@rev", r(11))
                    .Parameters.AddWithValue("@aut", r(12))
                End With

            Case "B.2"
                initialQuery &= "INSERT INTO B.2 (Secretaria,Dirección,No_Nomina,Situación,LugarComisión,DíasComicionado,Observaciones,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@No_Nomina,@Situacion,@LugarComision,@DiasComicionado,@Observaciones,@CveCorteEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@No_Nomina", r(2))
                    .Parameters.AddWithValue("@Situacion", r(3))
                    .Parameters.AddWithValue("@LugarComision", r(4))
                    .Parameters.AddWithValue("@DiasComicionado", r(5))
                    .Parameters.AddWithValue("@Observaciones", r(6))
                    .Parameters.AddWithValue("@Sindicalizado ", r(7))
                    .Parameters.AddWithValue("@CveCorteEjer ", r(8))
                    .Parameters.AddWithValue("@elab", r(9))
                    .Parameters.AddWithValue("@rev", r(10))
                    .Parameters.AddWithValue("@aut", r(11))
                End With

            Case "B.3"
                initialQuery &= "INSERT INTO B.3 (Secretaria,Dirección,Turno,NúmerodeEmpleado,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@Turno,@NumerodeEmp,@CveCorteEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@Turno", r(2))
                    .Parameters.AddWithValue("@NumerodeEmp", r(3))
                    .Parameters.AddWithValue("@CveCorteEjer ", r(4))
                    .Parameters.AddWithValue("@elab", r(5))
                    .Parameters.AddWithValue("@rev", r(6))
                    .Parameters.AddWithValue("@aut", r(7))
                End With

            Case "B.4"
                initialQuery &= "INSERT INTO B.4 
(Secretaria,Dirección,No_Nomina,NombreCompleto,Clasificación,PersepciónMensual,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES 
(@sec,@dir,@No_Nom,@Nom_Comp,@Clasf,@Pers_Men,@Cve_Corte,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@No_Nom", r(2))
                    .Parameters.AddWithValue("@Nom_Comp", r(3))
                    .Parameters.AddWithValue("@Clasf", r(4))
                    .Parameters.AddWithValue("@Pers_Men", r(5))
                    .Parameters.AddWithValue("@Cve_Corte", r(6))
                    .Parameters.AddWithValue("@elab", r(7))
                    .Parameters.AddWithValue("@rev", r(8))
                    .Parameters.AddWithValue("@aut", r(9))
                End With



            Case "C.2"
                initialQuery &= "INSERT INTO C.2
 (Secretaria, Dirección, No. Inventario, Descripción, Marca, Modelo, Placa, No. Serie, No. Nomina Resguardante, Condiciones, Tipo Combustible, Capacidad Combustible, Estación Asignada, No. Poliza Seguro, Cobertura Poliza, CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES
(@sec,@dir, @No.Inv, @des, @mar, @mod, @pla, @No.Ser, @No.NomRes, @con, @TipCom, @CapCom, @EstAsi, @No.PolSeg, @CobPol, @CveCorEje, @ela, @rev, @aut)"

                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@No.Inv", r(2))
                    .Parameters.AddWithValue("@des", r(3))
                    .Parameters.AddWithValue("@mar", r(4))
                    .Parameters.AddWithValue("@mod", r(5))
                    .Parameters.AddWithValue("@pla", r(6))
                    .Parameters.AddWithValue("@No.Ser", r(7))
                    .Parameters.AddWithValue("@No.NomRes", r(8))
                    .Parameters.AddWithValue("@con", r(9))
                    .Parameters.AddWithValue("@TipCom", r(10))
                    .Parameters.AddWithValue("@CapCom", r(11))
                    .Parameters.AddWithValue("@EstAsi", r(12))
                    .Parameters.AddWithValue("@No.PolSeg", r(13))
                    .Parameters.AddWithValue("@CobPol", r(14))
                    .Parameters.AddWithValue("@CveCorEje", r(15))
                    .Parameters.AddWithValue("@ela", r(16))
                    .Parameters.AddWithValue("@rev", r(17))
                    .Parameters.AddWithValue("@aut", r(18))
                End With

            Case "C.3"
                initialQuery &= "INSERT INTO C.3
 (Secretaria, Dirección, Titulo, Numero Ejemplares, CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES
(@sec,@dir, @tit,@NumEje, @CveCorEje, @ela, @rev, @aut)"

                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@tit", r(2))
                    .Parameters.AddWithValue("@NumEje", r(3))
                    .Parameters.AddWithValue("@CveCorEje", r(4))
                    .Parameters.AddWithValue("@ela", r(5))
                    .Parameters.AddWithValue("@rev", r(6))
                    .Parameters.AddWithValue("@aut", r(7))
                End With

            Case "C.3.1"
                initialQuery &= "INSERT INTO C.3.1
 (Secretaria, Dirección, Titulo, Numero Ejemplares, CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES
(@sec,@dir, @tit,@NumEje, @CveCorEje, @ela, @rev, @aut)"

                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@tit", r(2))
                    .Parameters.AddWithValue("@NumEje", r(3))
                    .Parameters.AddWithValue("@CveCorEje", r(4))
                    .Parameters.AddWithValue("@ela", r(5))
                    .Parameters.AddWithValue("@rev", r(6))
                    .Parameters.AddWithValue("@aut", r(7))
                End With

            Case "C.7"
                initialQuery &= "INSERT INTO C.7
(Secretaria,Dirección,NominaResgardante,NombreRresguardante,TipoArma,Calibre,NúmeroSerie,Origen,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Secretaria,Dirección,@NominaResgardante,@NombreRresguardante,@TipoArma,@Calibre,@NúmeroSerie,@Origen,@CveCorteEjer,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@NominaResgardante", r(2))
                    .Parameters.AddWithValue("@NombreResguardante", r(3))
                    .Parameters.AddWithValue("@TipoArma", r(4))
                    .Parameters.AddWithValue("@Calibre", r(5))
                    .Parameters.AddWithValue("@NúmeroSerie", r(6))
                    .Parameters.AddWithValue("@Origen", r(7))
                    .Parameters.AddWithValue("@CveCorteEjer", r(8))
                    .Parameters.AddWithValue("@Elaboró", r(9))
                    .Parameters.AddWithValue("@Revisó", r(10))
                    .Parameters.AddWithValue("@Autorizó", r(11))
                End With

            Case "C.8"
                initialQuery &= "INSERT INTO C.8
(Secretaria,Dirección,TipoFormato,Cantidad,NombreEvento,Fecha,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Secretaria,Dirección,@TipoFormato,@Cantidad,@NombreEvento,@Fecha,@CveCorteEjer,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@TipoFormato", r(2))
                    .Parameters.AddWithValue("@Cantidad", r(3))
                    .Parameters.AddWithValue("@NombreEvento", r(4))
                    .Parameters.AddWithValue("@Fecha", r(5))
                    .Parameters.AddWithValue("@CveCorteEjer", r(6))
                    .Parameters.AddWithValue("@Elaboró", r(7))
                    .Parameters.AddWithValue("@Revisó", r(8))
                    .Parameters.AddWithValue("@Autorizó", r(9))
                End With

            Case "C.9"
                initialQuery &= "INSERT INTO C.9
(Secretaria,Dirección,Clasificación,NúmeroEjemplares,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Secretaria,Dirección,@Clasificación,@NúmeroEjemplares@CveCorteEjer,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@Clasificación", r(2))
                    .Parameters.AddWithValue("@NúmeroEjemplares", r(3))
                    .Parameters.AddWithValue("@CveCorteEjer", r(4))
                    .Parameters.AddWithValue("@Elaboró", r(5))
                    .Parameters.AddWithValue("@Revisó", r(6))
                    .Parameters.AddWithValue("@Autorizó", r(7))
                End With

            Case "D.2"
                initialQuery &= "INSERT INTO D.2
(Secretaria,Dirección,No.Contrato,Descripción,ContatistaAsignado,MontoObra,MontoEjercicio,ModalidadContrato,UbicaExpediente,RecursoUtilizado,PorcentajeAvance,Metas,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@NoCon,@Desc,@Cont,@ContAsig,@MObra,@MEje,@ModCon,@UbiExp,@RecUt,@PorAvan,@Metas,@CveCortEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@NoCon ", r(2))
                    .Parameters.AddWithValue("@Desc", r(3))
                    .Parameters.AddWithValue("@Cont", r(4))
                    .Parameters.AddWithValue("@ContAsig", r(5))
                    .Parameters.AddWithValue("@MObra", r(6))
                    .Parameters.AddWithValue("@MEje", r(7))
                    .Parameters.AddWithValue("@ModCon", r(8))
                    .Parameters.AddWithValue("@UbiExp", r(9))
                    .Parameters.AddWithValue("@RecUt", r(10))
                    .Parameters.AddWithValue("@PorAvan", r(11))
                    .Parameters.AddWithValue("@Metas", r(12))
                    .Parameters.AddWithValue("@CveCortEjer", r(13))
                    .Parameters.AddWithValue("@elab", r(14))
                    .Parameters.AddWithValue("@rev", r(15))
                    .Parameters.AddWithValue("@aut", r(16))
                End With

            Case "D.3"
                initialQuery &= "INSERT INTO D.3
(Secretaria,Dirección, Indicador, NombreProg, Descripción, PorcentajeCump, CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@Ind,@NomPro,@Desc,@PorCump,@CveCorteEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@Ind ", r(2))
                    .Parameters.AddWithValue("@NomPro", r(3))
                    .Parameters.AddWithValue("@Desc", r(4))
                    .Parameters.AddWithValue("@PorCump", r(5))
                    .Parameters.AddWithValue("@CveCorteEjer", r(6))
                    .Parameters.AddWithValue("@elab", r(7))
                    .Parameters.AddWithValue("@rev", r(8))
                    .Parameters.AddWithValue("@aut", r(9))
                End With

            Case "D.4"
                initialQuery &= "INSERT INTO D.4 (Secretaria,Dirección,No,NúmContrato,Contratación,Modalidad,Monto,FuenteFinan,FechaConclusión,Observaciones,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@No,@NumCont,@Cont,@Mod,@Mont,@FuenFin,@FechaCon,@Obser,@CveCorteEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@No", r(2))
                    .Parameters.AddWithValue("@NumCont", r(3))
                    .Parameters.AddWithValue("@Cont", r(4))
                    .Parameters.AddWithValue("@Mod", r(5))
                    .Parameters.AddWithValue("@Mont", r(6))
                    .Parameters.AddWithValue("@FuenFin", r(7))
                    .Parameters.AddWithValue("@FechaCon", r(8))
                    .Parameters.AddWithValue("@Obser", r(9))
                    .Parameters.AddWithValue("@CveCorteEjer", r(10))
                    .Parameters.AddWithValue("@elab", r(11))
                    .Parameters.AddWithValue("@rev", r(12))
                    .Parameters.AddWithValue("@aut", r(13))
                End With

            Case "D.5"
                initialQuery &= "INSERT INTO D.5 (Secretaria,Dirección,No,NúmContrato,Contratación,Modalidad,Monto,FuenteFinan,FechaConclusión,Observaciones,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@No,@NumCont,@Cont,@Mod,@Mont,@FuenFin,@FechaCon,@Obser,@CveCorteEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@No", r(2))
                    .Parameters.AddWithValue("@NumCont", r(3))
                    .Parameters.AddWithValue("@Cont", r(4))
                    .Parameters.AddWithValue("@Mod", r(5))
                    .Parameters.AddWithValue("@Mont", r(6))
                    .Parameters.AddWithValue("@FuenFin", r(7))
                    .Parameters.AddWithValue("@FechaCon", r(8))
                    .Parameters.AddWithValue("@Obser", r(9))
                    .Parameters.AddWithValue("@CveCorteEjer", r(10))
                    .Parameters.AddWithValue("@elab", r(11))
                    .Parameters.AddWithValue("@rev", r(12))
                    .Parameters.AddWithValue("@aut", r(13))
                End With

        End Select
    End Function

End Class
