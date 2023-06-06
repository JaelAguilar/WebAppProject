Imports System.Data.SqlClient
Imports System.Security.Cryptography

Partial Class WebForm1
    Private Function GenerateSQLCommand(sql As SqlCommand, r As DataRow)
        Dim initialQuery As String = String.Empty
        Dim databaseName = importTableSelector.SelectedValue
        Debug.WriteLine("DATABASE = " + databaseName)

        Select Case databaseName
            Case "A.1"
                Debug.Write("Case A.1")
                initialQuery &= "INSERT INTO dbo.[A.1] (Secretaria,Dirección,I_PreEst,I_PreMod,I_PreDev,I_PreRec,E_PreOrigApro,E_1A_AmpPres,E_2A_AmpPres,E_3A_AmpPres,E_Tot_Amp,E_PreModif,E_PreComp,E_PreDev,E_PreEjer,E_PreErog,E_PreCons,E_PrePorEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@IpreEst,@IpreMod,@IpreDev,@IpreRec,@EorigApro,@1ampPres,@2ampPres,@3ampPres,@EtotAmp,@EpreModif,@EpreComp,@EpreDev,@EpreEjer,@EpreErog,@EpreCons,@EprePorEjer,@elab,@rev,@aut)"
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
                    .Parameters.AddWithValue("@elab", r(18))
                    .Parameters.AddWithValue("@rev", r(19))
                    .Parameters.AddWithValue("@aut", r(20))
                End With

            Case "A.1.1"
                initialQuery &= "INSERT INTO dbo.[A.1.1]
(Secretaria, Dirección,Impuestos,Cuot_ApSS,Cont_Mej,Derechos,Productos,Aprov,Ing_Vta_Bs,Part_Apor,Tras_Sub_Oayu,Ing_Finan,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@sec,@Direccion,@Impuestos,@Cuot_ApSS,@Cont_Mej,@Derechos,@Productos,@Aprov,@Ing_Vta_Bs,@Part_Apor,@Tras_Sub_Oayu,@Ing_Finan,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@Direccion", r(1))
                    .Parameters.AddWithValue("@Impuestos", r(2))
                    .Parameters.AddWithValue("@Cuot_ApSS", r(3))
                    .Parameters.AddWithValue("@Cont_Mej", r(4))
                    .Parameters.AddWithValue("@Derechos", r(5))
                    .Parameters.AddWithValue("@Productos", r(6))
                    .Parameters.AddWithValue("@Aprov", r(7))
                    .Parameters.AddWithValue("@Ing_Vta_Bs", r(8))
                    .Parameters.AddWithValue("@Part_Apor", r(9))
                    .Parameters.AddWithValue("@Tras_Sub_Oayu", r(10))
                    .Parameters.AddWithValue("@Ing_Finan", r(11))
                    .Parameters.AddWithValue("@Elaboró", r(12))
                    .Parameters.AddWithValue("@Revisó", r(13))
                    .Parameters.AddWithValue("@Autorizó", r(14))
                End With

            Case "A.1.2"
                initialQuery &= "INSERT INTO dbo.[A.1.2]
(Secretaria, Dirección,Serv_Per,Mat_Sum,Serv_Gen,Tras_Sub_Oayu,Bien_Mu_Inm_Inta,Iver_Pub,Inver_Fin_OP,Part_Apor,Dedua_Pub,FechaCorte,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Secretaria,@Dirección,@Serv_Per,@Mat_Sum,@Serv_Gen,@Tras_Sub_Oayu,@Bien_Mu_Inm_Inta,@Iver_Pub,@Inver_Fin_OP,@Part_Apor,@Dedua_Pub,@FechaCorte,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@Serv_Per", r(2))
                    .Parameters.AddWithValue("@Mat_Sum", r(3))
                    .Parameters.AddWithValue("@Serv_Gen", r(4))
                    .Parameters.AddWithValue("@Tras_Sub_Oayu", r(5))
                    .Parameters.AddWithValue("@Bien_Mu_Inm_Inta", r(6))
                    .Parameters.AddWithValue("@Iver_Pub", r(7))
                    .Parameters.AddWithValue("@Inver_Fin_OP", r(8))
                    .Parameters.AddWithValue("@Part_Apor", r(9))
                    .Parameters.AddWithValue("@Dedua_Pub", r(10))
                    .Parameters.AddWithValue("@FechaCorte", r(11))
                    .Parameters.AddWithValue("@Elaboró", r(12))
                    .Parameters.AddWithValue("@Revisó", r(13))
                    .Parameters.AddWithValue("@Autorizó", r(14))
                End With

            Case "A.1.3"
                initialQuery &= "INSERT INTO dbo.[A.1.3]
(Secretaria,Dirección,Clave,Nombre,Presup_Auto,Porcentaje,FechaCorte,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Secretaria,@Dirección,@Clave,@Nombre,@Presup_Auto,@Porcentaje,@FechaCorte,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@Clave", r(2))
                    .Parameters.AddWithValue("@Nombre", r(3))
                    .Parameters.AddWithValue("@Presup_Auto", r(4))
                    .Parameters.AddWithValue("@Porcentaje", r(5))
                    .Parameters.AddWithValue("@FechaCorte", r(6))
                    .Parameters.AddWithValue("@Elaboró", r(7))
                    .Parameters.AddWithValue("@Revisó", r(8))
                    .Parameters.AddWithValue("@Autorizó", r(9))
                End With

            Case "A.1.4"
                initialQuery &= "INSERT INTO dbo.[A.1.4]
(Secretaria,Dirección,Presup_Auto,1a_AmpPre,2a_AmpPre,3a_AmpPre,Total_Amp,Pre_Modificado,Pre_Comprometido,Pre_Devengado,Pre_Ejercicio,Pre_Erogado,Pre_Consuido,Pre_PorEjercer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Secretaria,@Dirección,@Presup_Auto,@1a_AmpPre,@2a_AmpPre,@3a_AmpPre,@Total_Amp,@Pre_Modificado,@Pre_Comprometido,@Pre_Devengado,@Pre_Ejercicio,@Pre_Erogado,@Pre_Consuido,@Pre_PorEjercer,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@Presup_Auto", r(2))
                    .Parameters.AddWithValue("@1a_AmpPre", r(3))
                    .Parameters.AddWithValue("@2a_AmpPre", r(4))
                    .Parameters.AddWithValue("@3a_AmpPre", r(5))
                    .Parameters.AddWithValue("@Total_Amp", r(6))
                    .Parameters.AddWithValue("@Pre_Modificado", r(7))
                    .Parameters.AddWithValue("@Pre_Comprometido", r(8))
                    .Parameters.AddWithValue("@Pre_Devengado", r(9))
                    .Parameters.AddWithValue("@Pre_Ejercicio", r(10))
                    .Parameters.AddWithValue("@Pre_Erogado", r(11))
                    .Parameters.AddWithValue("@Pre_Consuido", r(12))
                    .Parameters.AddWithValue("@Pre_PorEjercer", r(13))
                    .Parameters.AddWithValue("@Elaboró", r(14))
                    .Parameters.AddWithValue("@Revisó", r(15))
                    .Parameters.AddWithValue("@Autorizó", r(16))
                End With

            Case "A.2"
                initialQuery &= "INSERT INTO dbo.[A.2]
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
                initialQuery &= "INSERT INTO dbo.[A.3]
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
                initialQuery &= "INSERT INTO dbo.[A.4]
(Secretaria,Dirección,Nom_Secretaria,Titular,Monto,FechaCorte,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Sec,@Dir,@NomSec,@Titular,@Monto,@FechaCorte,@Elab,@Rev,@Aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Sec", r(0))
                    .Parameters.AddWithValue("@Dir", r(1))
                    .Parameters.AddWithValue("@NomSec", r(2))
                    .Parameters.AddWithValue("@Titular", r(3))
                    .Parameters.AddWithValue("@Monto", r(4))
                    .Parameters.AddWithValue("@FechaCorte", r(5))
                    .Parameters.AddWithValue("@Elab", r(6))
                    .Parameters.AddWithValue("@Rev", r(7))
                    .Parameters.AddWithValue("@Aut", r(8))
                End With

            Case "A.4.1a"
                initialQuery &= "INSERT INTO dbo.[A.4.1a]
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
                initialQuery &= "INSERT INTO dbo.[A.4.1b]
(Secretaria,Dirección,Fecha,Documento,Fecha_Doc,Proveedor,Concepto,Importe,FechaCorte,Elaboró,Revisó,Autorizó)"
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
                initialQuery &= "INSERT INTO dbo.[A.5]
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
                initialQuery &= "INSERT INTO dbo.[A.5.1]
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
                    .Parameters.AddWithValue("@SaldoSL", r(5))
                    .Parameters.AddWithValue("@SaldoSECB", r(6))
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

            Case "A.5.1.1"
                initialQuery &= "INSERT INTO dbo.[A.5.1.1] (Secretaria, Dirección,CCCH,ClaveEjercicio, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@CCCH,@ClaveEjercicio,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("CCCH", r(2))
                    .Parameters.AddWithValue("ClaveEjercicio", r(3))
                    .Parameters.AddWithValue("Elaboro", r(4))
                    .Parameters.AddWithValue("reviso", r(5))
                    .Parameters.AddWithValue("autorizo", r(6))
                End With

            Case "A.5.2"
                initialQuery &= "INSERT INTO dbo.[A.5.2] (Secretaria, Dirección,Nom_Inst,Num_Cuenta,Cta_Contable,Saldo_SL,Saldo_SECB,Tipo_Inver,Vencimineto,Frima_R1,Frima_R2,Frima_R3,Frima_R4,Cta_Individual,Cta_Mancomunada,Cta_Indistinta,ClaveEjercicio, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@Nom_Inst,@Num_Cuenta,@Cta_Contable,@Saldo_SL,@Saldo_SECB,@Tipo_Inver,@Vencimineto,@Frima_R1,@Frima_R2,@Frima_R3,@Frima_R4,@Cta_Individual,@Cta_Mancomunada,@Cta_Indistinta,@ClaveEjercicio,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("Nom_Inst", r(2))
                    .Parameters.AddWithValue("Num_Cuenta", r(3))
                    .Parameters.AddWithValue("Cta_Contable", r(4))
                    .Parameters.AddWithValue("Saldo_SL", r(5))
                    .Parameters.AddWithValue("Saldo_SECB", r(6))
                    .Parameters.AddWithValue("Tipo_Inver", r(7))
                    .Parameters.AddWithValue("Vencimineto", r(8))
                    .Parameters.AddWithValue("Frima_R1", r(9))
                    .Parameters.AddWithValue("Frima_R2", r(10))
                    .Parameters.AddWithValue("Frima_R3", r(11))
                    .Parameters.AddWithValue("Frima_R4", r(12))
                    .Parameters.AddWithValue("Cta_Individual", r(13))
                    .Parameters.AddWithValue("Cta_Mancomunada", r(14))
                    .Parameters.AddWithValue("Cta_Indistinta", r(15))
                    .Parameters.AddWithValue("ClaveEjercicio", r(16))
                    .Parameters.AddWithValue("Elaboro", r(17))
                    .Parameters.AddWithValue("reviso", r(18))
                    .Parameters.AddWithValue("autorizo", r(19))
                End With

            Case "A.5.2.1"
                initialQuery &= "INSERT INTO dbo.[A.5.2.1] (Secretaria, Dirección,CCI,ClaveEjercicio, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@CCI,@ClaveEjercicio,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("CCI", r(2))
                    .Parameters.AddWithValue("ClaveEjercicio", r(3))
                    .Parameters.AddWithValue("Elaboro", r(4))
                    .Parameters.AddWithValue("reviso", r(5))
                    .Parameters.AddWithValue("autorizo", r(6))
                End With

            Case "A.6"
                initialQuery &= "INSERT INTO [A.6] (Secretaria,Dirección,Num_Documento,Nom_Deudor,Fech_Adeudo,Importe_Tot,Saldo,Vencimiento,Concepto,ClaveEjercicio, Elaboró,Revisó,Autorizó)"
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
                initialQuery &= "INSERT INTO dbo.[A.6.2] (Secretaria,Dirección,RSSI,ClaveEjercicio, Elaboró,Revisó,Autorizó)"
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
                initialQuery &= "INSERT INTO dbo.[A.7] (Secretaria,Dirección,Num_Documento,Nom_Acreador,Saldo,ClaveEjercicio, Elaboró,Revisó,Autorizó)"
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
                initialQuery &= "INSERT INTO dbo.[A.7.1] (Secretaria,Dirección,NomProveedor,Saldo,ClaveEjercicio,Elaboró,Revisó,Autorizó)"
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
                initialQuery &= "INSERT INTO dbo.[A.7.2] (Secretaria,Dirección,NomProveedor,Saldo,CveCorteEjer,Elaboró,Revisó,Autorizó)"
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
                initialQuery &= "INSERT INTO dbo.[A.7.3] (Secretaria,Dirección,NoDocumento,NomAcredor,Fecha,Saldo,Vencimiento,Concepto,CveCorteEjer,Elaboró,Revisó,Autorizó)"
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
                initialQuery &= "INSERT INTO dbo.[A.8] (Secretaria,Dirección,Fecha,NoCtaBanco,Institución,NoCheque,NomBenif,Importe,CveCorteEjer,Elaboró,Revisó,Autorizo)"
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
                initialQuery &= "INSERT INTO dbo.[A.9] (Secretaria,Dirección,NoPoliza,NomAfianzadora,NomDeudor,Monto,ConceptoFianza,CveCorteEjer,Elaboró,Revisó,Autorizo)"
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
                initialQuery &= "INSERT INTO dbo.[A.10] (Secretaria,Dirección,No,EjercicioFiscal,ActaEAESNL,NoLegajos,NoDiscos,Estatus,Observaciones,Responsable,CveCorteEjer,Elaboró,Revisó,Autorizo)"
                initialQuery &= "VALUES (@sec,@dir,@No,@EjercicioFiscal,@ActaEAESNL,@NoLegajos,@NoDiscos,@Estatus,@Observaciones,@Responsable,@CveCorteEjer,@elab,@rev,@aut)"
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
                    .Parameters.AddWithValue("@Responsable", r(9))
                    .Parameters.AddWithValue("@CveCorteEjer", r(10))
                    .Parameters.AddWithValue("@elab ", r(11))
                    .Parameters.AddWithValue("@rev ", r(12))
                    .Parameters.AddWithValue("@aut ", r(13))
                End With

            Case "B.1"
                initialQuery &= "INSERT INTO dbo.[B.1] (Secretaria,Dirección,No_Nomina,Puesto,NombreCompleto,Sueldo,Vigencia,Sindicalizado,RegimenEmpleado,CveCorteEjer,Elaboró,Revisó,Autorizó)"
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
                initialQuery &= "INSERT INTO dbo.[B.2] (Secretaria,Dirección,No_Nomina,Situación,LugarComisión,DíasComicionado,Observaciones,CveCorteEjer,Elaboró,Revisó,Autorizó)"
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
                    .Parameters.AddWithValue("@CveCorteEjer ", r(7))
                    .Parameters.AddWithValue("@elab", r(8))
                    .Parameters.AddWithValue("@rev", r(9))
                    .Parameters.AddWithValue("@aut", r(10))
                End With

            Case "B.3"
                initialQuery &= "INSERT INTO dbo.[B.3] (Secretaria,Dirección,Turno,NúmerodeEmpleado,CveCorteEjer,Elaboró,Revisó,Autorizó)"
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
                initialQuery &= "INSERT INTO dbo.[B.4]
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
                initialQuery &= "INSERT INTO dbo.[C.2]
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
                initialQuery &= "INSERT INTO dbo.[C.3]
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
                initialQuery &= "INSERT INTO dbo.[C.3.1]
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

            Case "C.4"
                initialQuery &= "INSERT INTO dbo.[C.4] (Secretaria,Dirección,Cantidad,NombreFormato,FolioInicial,FolioFinal,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@cant,@NomForm,@FolIni,@FolFin,@CveCortEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@cant", r(2))
                    .Parameters.AddWithValue("@NomForm", r(3))
                    .Parameters.AddWithValue("@FolIni", r(4))
                    .Parameters.AddWithValue("@FolFin", r(5))
                    .Parameters.AddWithValue("@CveCortEjer", r(6))
                    .Parameters.AddWithValue("@elab", r(7))
                    .Parameters.AddWithValue("@rev", r(8))
                    .Parameters.AddWithValue("@aut", r(9))
                End With

            Case "C.5"
                initialQuery &= "INSERT INTO dbo.[C.5] (Secretaria,Dirección,Código,Descripción,Cantidad,Condiciones,Observaciones,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@Cod,@Desc,@Cant,@Cond,@Obser,@CveCortEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@Cod", r(2))
                    .Parameters.AddWithValue("@Desc", r(3))
                    .Parameters.AddWithValue("@Cant", r(4))
                    .Parameters.AddWithValue("@Cond", r(2))
                    .Parameters.AddWithValue("@Obser", r(5))
                    .Parameters.AddWithValue("@CveCortEjer", r(6))
                    .Parameters.AddWithValue("@elab", r(7))
                    .Parameters.AddWithValue("@rev", r(8))
                    .Parameters.AddWithValue("@aut", r(9))
                End With

            Case "C.6"
                initialQuery &= "INSERT INTO dbo.[C.6] (Secretaria,Dirección,Código,Descripción,No.Inventario,NombrePropietario,Observaciones,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@Cod,@Desc,@NoInv,@NomPro,@Obser,@CveCortEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@Cod", r(2))
                    .Parameters.AddWithValue("@Desc", r(3))
                    .Parameters.AddWithValue("@NoInv", r(4))
                    .Parameters.AddWithValue("@NomPro", r(2))
                    .Parameters.AddWithValue("@Obser", r(5))
                    .Parameters.AddWithValue("@CveCortEjer", r(6))
                    .Parameters.AddWithValue("@elab", r(7))
                    .Parameters.AddWithValue("@rev", r(8))
                    .Parameters.AddWithValue("@aut", r(9))
                End With



            Case "C.7"
                initialQuery &= "INSERT INTO dbo.[C.7]
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
                initialQuery &= "INSERT INTO dbo.[C.8]
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
                initialQuery &= "INSERT INTO dbo[C.9]
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

            Case "C.10"
                initialQuery &= "INSERT INTO dbo.[C.10]
(Secretaria, Dirección,No.Inventario,EqCan,Nombre,FierroChip,Descripción,FechaNacimineto,FcehaAdquisición,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Secretaria,@Dirección,@No.Inventario,@EqCan,@Nombre,@FierroChip,@Descripción,@FechaNacimineto,@FcehaAdquisición,@CveCorteEjer,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@No.Inventario", r(2))
                    .Parameters.AddWithValue("@EqCan", r(3))
                    .Parameters.AddWithValue("@Nombre", r(4))
                    .Parameters.AddWithValue("@FierroChip", r(5))
                    .Parameters.AddWithValue("@Descripción", r(6))
                    .Parameters.AddWithValue("@FechaNacimineto", r(7))
                    .Parameters.AddWithValue("@FcehaAdquisición", r(8))
                    .Parameters.AddWithValue("@CveCorteEjer", r(9))
                    .Parameters.AddWithValue("@Elaboró", r(10))
                    .Parameters.AddWithValue("@Revisó", r(11))
                    .Parameters.AddWithValue("@Autorizó", r(12))
                End With

            Case "C.11"
                initialQuery &= "INSERT INTO dbo.[C.11] (Secretaria, Dirección,No.Expediente,Uso,Ubicación,Superficie,Estatus,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@No.Expediente,@Uso,@Ubicación,@Superficie,@Estatus,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("No.Expediente", r(2))
                    .Parameters.AddWithValue("Uso", r(3))
                    .Parameters.AddWithValue("Ubicación", r(4))
                    .Parameters.AddWithValue("Superficie", r(5))
                    .Parameters.AddWithValue("Estatus", r(6))
                    .Parameters.AddWithValue("CveCorteEjer", r(7))
                    .Parameters.AddWithValue("Elaboro", r(8))
                    .Parameters.AddWithValue("reviso", r(9))
                    .Parameters.AddWithValue("autorizo", r(10))
                End With

            Case "C.11.1"
                initialQuery &= "INSERT INTO dbo.[C.11.1] (Secretaria, Dirección,No.Expediente,Uso,Ubicación,Superficie,Estatus,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@No.Expediente,@Uso,@Ubicación,@Superficie,@Estatus,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("No.Expediente", r(2))
                    .Parameters.AddWithValue("Uso", r(3))
                    .Parameters.AddWithValue("Ubicación", r(4))
                    .Parameters.AddWithValue("Superficie", r(5))
                    .Parameters.AddWithValue("Estatus", r(6))
                    .Parameters.AddWithValue("CveCorteEjer", r(7))
                    .Parameters.AddWithValue("Elaboro", r(8))
                    .Parameters.AddWithValue("reviso", r(9))
                    .Parameters.AddWithValue("autorizo", r(10))
                End With

            Case "C.11.2"
                initialQuery &= "INSERT INTO dbo.[C.11.2] (Secretaria, Dirección,No.Expediente,Uso,Ubicación,NombreComodatario,Vigencia,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@No.Expediente,@Uso,@Ubicación,@NombreComodatario,@Vigencia,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("No.Expediente", r(2))
                    .Parameters.AddWithValue("Uso", r(3))
                    .Parameters.AddWithValue("Ubicación", r(4))
                    .Parameters.AddWithValue("NombreComodatario", r(5))
                    .Parameters.AddWithValue("Vigencia", r(6))
                    .Parameters.AddWithValue("CveCorteEjer", r(7))
                    .Parameters.AddWithValue("Elaboro", r(8))
                    .Parameters.AddWithValue("reviso", r(9))
                    .Parameters.AddWithValue("autorizo", r(10))
                End With

            Case "D.1"
                initialQuery &= "INSERT INTO dbo.[D.1] (Secretaria, Dirección,NombreProveedor,Especialidad,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@NombreProveedor,@Especialidad,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("NombreProveedor", r(2))
                    .Parameters.AddWithValue("Especialidad", r(3))
                    .Parameters.AddWithValue("CveCorteEjer", r(4))
                    .Parameters.AddWithValue("Elaboro", r(5))
                    .Parameters.AddWithValue("reviso", r(6))
                    .Parameters.AddWithValue("autorizo", r(7))
                End With

            Case "D.2"
                initialQuery &= "INSERT INTO dbo.[D.2]
(Secretaria,Dirección,NoContrato,Descripción,ContatistaAsignado,MontoObra,MontoEjercicio,ModalidadContrato,UbicaExpediente,RecursoUtilizado,PorcentajeAvance,Metas,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES (@sec,@dir,@NoCon,@Desc,@ContAsig,@MObra,@MEje,@ModCon,@UbiExp,@RecUt,@PorAvan,@Metas,@CveCortEjer,@elab,@rev,@aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@NoCon ", r(2))
                    .Parameters.AddWithValue("@Desc", r(3))
                    .Parameters.AddWithValue("@ContAsig", r(4))
                    .Parameters.AddWithValue("@MObra", r(5))
                    .Parameters.AddWithValue("@MEje", r(6))
                    .Parameters.AddWithValue("@ModCon", r(7))
                    .Parameters.AddWithValue("@UbiExp", r(8))
                    .Parameters.AddWithValue("@RecUt", r(9))
                    .Parameters.AddWithValue("@PorAvan", r(10))
                    .Parameters.AddWithValue("@Metas", r(11))
                    .Parameters.AddWithValue("@CveCortEjer", r(12))
                    .Parameters.AddWithValue("@elab", r(13))
                    .Parameters.AddWithValue("@rev", r(14))
                    .Parameters.AddWithValue("@aut", r(15))
                End With

            Case "D.3"
                initialQuery &= "INSERT INTO dbo.[D.3]
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
                initialQuery &= "INSERT INTO dbo.[D.4] (Secretaria,Dirección,No,NúmContrato,Contratación,Modalidad,Monto,FuenteFinan,FechaConclusión,Observaciones,CveCorteEjer,Elaboró,Revisó,Autorizó)"
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
                initialQuery &= "INSERT INTO dbo.[D.5] (Secretaria,Dirección,No,NúmContrato,Contratación,Modalidad,Monto,FuenteFinan,FechaConclusión,Observaciones,CveCorteEjer,Elaboró,Revisó,Autorizó)"
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

            Case "D.6"
                initialQuery &= "INSERT INTO dbo.[D.6] (Secretaria, Dirección,No,NúmContrato,Contratación,Modalidad,Monto,FuenteFinan,FechaConclusión,Observaciones,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@No,@NúmContrato,@Contratación,@Modalidad,@Monto,@FuenteFinan,@FechaConclusión,@Observaciones,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("No", r(2))
                    .Parameters.AddWithValue("NúmContrato", r(3))
                    .Parameters.AddWithValue("Contratación", r(4))
                    .Parameters.AddWithValue("Modalidad", r(5))
                    .Parameters.AddWithValue("Monto", r(6))
                    .Parameters.AddWithValue("FuenteFinan", r(7))
                    .Parameters.AddWithValue("FechaConclusión", r(8))
                    .Parameters.AddWithValue("Observaciones", r(9))
                    .Parameters.AddWithValue("CveCorteEjer", r(10))
                    .Parameters.AddWithValue("Elaboro", r(11))
                    .Parameters.AddWithValue("reviso", r(12))
                    .Parameters.AddWithValue("autorizo", r(13))
                End With

            Case "D.7"
                initialQuery &= "INSERT INTO dbo.[D.7] (Secretaria, Dirección,NoContrato,Vigencia,UbicaciónObra,UbicaciónExp,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@NoContrato,@Vigencia,@UbicaciónObra,@UbicaciónExp,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("NoContrato", r(2))
                    .Parameters.AddWithValue("Vigencia", r(3))
                    .Parameters.AddWithValue("UbicaciónObra", r(4))
                    .Parameters.AddWithValue("UbicaciónExp", r(5))
                    .Parameters.AddWithValue("CveCorteEjer ", r(6))
                    .Parameters.AddWithValue("Elaboro", r(7))
                    .Parameters.AddWithValue("reviso", r(8))
                    .Parameters.AddWithValue("autorizo", r(9))
                End With

            Case "D.8"
                initialQuery &= "INSERT INTO dbo.[D.8] (Secretaria, Dirección,NoConsecutivo,FechaFormación,NúmeroIntegrantes,NúmeroContrato,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@NoConsecutivo,@FechaFormación,@NúmeroIntegrantes,@NúmeroContrato,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("NoConsecutivo", r(2))
                    .Parameters.AddWithValue("FechaFormación", r(3))
                    .Parameters.AddWithValue("NúmeroIntegrantes", r(4))
                    .Parameters.AddWithValue("NúmeroContrato", r(5))
                    .Parameters.AddWithValue("CveCorteEjer", r(6))
                    .Parameters.AddWithValue("Elaboro", r(7))
                    .Parameters.AddWithValue("reviso", r(8))
                    .Parameters.AddWithValue("autorizo", r(9))
                End With

            Case "E.1"
                initialQuery &= "INSERT INTO dbo.[E.1] (Secretaria, Dirección,NumExpediente,NumJusgado,Demandante,AutoridadResp,Demandado,EstadoProcesal,ConceptoDemanda,Observaciones,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@NumExpediente,@NumJusgado,@Demandante,@AutoridadResp,@Demandado,@EstadoProcesal,@ConceptoDemanda,@Observaciones,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("NumExpediente", r(2))
                    .Parameters.AddWithValue("NumJusgado", r(3))
                    .Parameters.AddWithValue("Demandante", r(4))
                    .Parameters.AddWithValue("AutoridadResp", r(5))
                    .Parameters.AddWithValue("Demandado", r(6))
                    .Parameters.AddWithValue("EstadoProcesal", r(7))
                    .Parameters.AddWithValue("ConceptoDemanda", r(8))
                    .Parameters.AddWithValue("Observaciones", r(9))
                    .Parameters.AddWithValue("CveCorteEjer", r(10))
                    .Parameters.AddWithValue("Elaboro", r(11))
                    .Parameters.AddWithValue("reviso", r(12))
                    .Parameters.AddWithValue("autorizo", r(13))
                End With

            Case "E.2"
                initialQuery &= "INSERT INTO dbo.[E.2] (Secretaria, Dirección,Tipo,FechaSuscrip,Duracción,Entidad,Descripción,Objeto,Situación,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@Tipo,@FechaSuscrip,@Duracción,@Entidad,@Descripción,@Objeto,@Situación,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("Tipo", r(2))
                    .Parameters.AddWithValue("FechaSuscrip", r(3))
                    .Parameters.AddWithValue("Duracción", r(4))
                    .Parameters.AddWithValue("Entidad", r(5))
                    .Parameters.AddWithValue("Descripción", r(6))
                    .Parameters.AddWithValue("Objeto", r(7))
                    .Parameters.AddWithValue("Situación", r(8))
                    .Parameters.AddWithValue("CveCorteEjer", r(9))
                    .Parameters.AddWithValue("Elaboro", r(10))
                    .Parameters.AddWithValue("reviso", r(11))
                    .Parameters.AddWithValue("autorizo", r(12))
                End With

            Case "E.3"
                initialQuery &= "INSERT INTO dbo.[E.3] (Secretaria, Dirección,Tipo,Personas,Duracción,FechaSuscrip,Descripción,Objeto,Situación,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@Tipo,@Personas,@Duracción,@FechaSuscrip,@Descripción,@Objeto,@Situación,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("Tipo", r(2))
                    .Parameters.AddWithValue("Personas", r(3))
                    .Parameters.AddWithValue("Duracción", r(4))
                    .Parameters.AddWithValue("FechaSuscrip", r(5))
                    .Parameters.AddWithValue("Descripción", r(6))
                    .Parameters.AddWithValue("Objeto", r(7))
                    .Parameters.AddWithValue("Situación", r(8))
                    .Parameters.AddWithValue("CveCorteEjer", r(9))
                    .Parameters.AddWithValue("Elaboro", r(10))
                    .Parameters.AddWithValue("reviso", r(11))
                    .Parameters.AddWithValue("autorizo", r(12))
                End With

            Case "E.4"
                initialQuery &= "INSERT INTO dbo.[E.4]
(Secretaria,Dirección,seccion,nombre,Domicilio,FechaNomb,observaciones,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Sec,@Dir,@Secc,@Nom,@Dom,@FechaNomb,@Observ,@CveCorteEjer,@Elab,@Rev,@Aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Sec", r(0))
                    .Parameters.AddWithValue("@Dir", r(1))
                    .Parameters.AddWithValue("@Secc", r(2))
                    .Parameters.AddWithValue("@Nom", r(3))
                    .Parameters.AddWithValue("@Dom", r(4))
                    .Parameters.AddWithValue("@FechaNomb", r(5))
                    .Parameters.AddWithValue("@Observ", r(6))
                    .Parameters.AddWithValue("@CveCorteEjer", r(7))
                    .Parameters.AddWithValue("@Elab", r(8))
                    .Parameters.AddWithValue("@Rev", r(9))
                    .Parameters.AddWithValue("@Aut", r(10))
                End With

            Case "E.5"
                initialQuery &= "INSERT INTO dbo.[E.5]
(Secretaria,Dirección,NombreContribuyente,Cantidad,Descripcion,Clasificacion,Motivo,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Sec,@Dir,@NomCont,@Cant,@Desc,@Clas,@Mot,@CveCorteEjer,@Elab,@Rev,@Aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Sec", r(0))
                    .Parameters.AddWithValue("@Dir", r(1))
                    .Parameters.AddWithValue("@NomCont", r(2))
                    .Parameters.AddWithValue("@Cant", r(3))
                    .Parameters.AddWithValue("@Desc", r(4))
                    .Parameters.AddWithValue("@Clas", r(5))
                    .Parameters.AddWithValue("@Mot", r(6))
                    .Parameters.AddWithValue("@CveCorteEjer", r(7))
                    .Parameters.AddWithValue("@Elab", r(8))
                    .Parameters.AddWithValue("@Rev", r(9))
                    .Parameters.AddWithValue("@Aut", r(10))
                End With

            Case "E.6"
                initialQuery &= "INSERT INTO dbo.[E.6]
(Secretaria,Dirección,NoExpediente,NombrePostor,Ubicacion,Superficie,AutorizóCabildo,AutorizóCongreso,TipoEnajena,NoDecreto,FechaDecreto,Observaciones,CveCorteEjer,Elaboró,Revisó,Autorizó)"
                initialQuery &= "VALUES
(@Sec,@Dir,@NoExp,@NoPos,@Ubic,@Super,@AutCab,@AutCon,@TipEna,@NoDec,@FechaDecreto,@Obse,@CveCorteEjer,@Elab,@Rev,@Aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Sec", r(0))
                    .Parameters.AddWithValue("@Dir", r(1))
                    .Parameters.AddWithValue("@NoExp", r(2))
                    .Parameters.AddWithValue("@NoPos", r(3))
                    .Parameters.AddWithValue("@Ubic", r(4))
                    .Parameters.AddWithValue("@Super", r(5))
                    .Parameters.AddWithValue("@AutCab", r(6))
                    .Parameters.AddWithValue("@AutCon", r(7))
                    .Parameters.AddWithValue("@TipEna", r(8))
                    .Parameters.AddWithValue("@NoDec", r(9))
                    .Parameters.AddWithValue("@FechaDecreto", r(10))
                    .Parameters.AddWithValue("@Obse", r(11))
                    .Parameters.AddWithValue("@CveCorteEjer", r(12))
                    .Parameters.AddWithValue("@Elab", r(13))
                    .Parameters.AddWithValue("@Rev", r(14))
                    .Parameters.AddWithValue("@Aut", r(15))
                End With

            Case "E.7"
                initialQuery = “INSERT INTO dbo.[E.7]
(Secretaria, Dirección, NoExpediente, Colonia, NoDecreto, LotesDesafec, LotesE, LotesSE, Observaciones, CveCorteEjer, Elaboró, Revisó, Autorizó) “
                initialQuery &= “VALUES
 (@sec, @dir, @noExp, @colonia, @noDecreto, @lotesDesafec, @lotesE, @lotesSE, @observaciones, @cveCorteEjer, @elaboro, @reviso, @autorizo)”
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue(“@sec”, r(0))
                    .Parameters.AddWithValue(“@dir”, r(1))
                    .Parameters.AddWithValue(“@noExp”, r(2))
                    .Parameters.AddWithValue(“@colonia”, r(3))
                    .Parameters.AddWithValue(“@noDecreto”, r(4))
                    .Parameters.AddWithValue(“@lotesDesafec”, r(5))
                    .Parameters.AddWithValue(“@lotesE”, r(6))
                    .Parameters.AddWithValue(“@lotesSE”, r(7))
                    .Parameters.AddWithValue(“@observaciones”, r(8))
                    .Parameters.AddWithValue(“@cveCorteEjer”, r(9))
                    .Parameters.AddWithValue(“@elaboro”, r(10))
                    .Parameters.AddWithValue(“@reviso”, r(11))
                    .Parameters.AddWithValue(“@autorizo”, r(12))
                End With

            Case "E.8"
                initialQuery = "INSERT INTO dbo.[E.8]
(Secretaria, Dirección, No, Libro, Período, Año, Ubicación, CveCorteEjer, Elaboró, Revisó, Autorizó) "
                initialQuery &= "VALUES
 (@sec, @dir, @no, @libro, @periodo, @anio, @ubicacion, @cveCorteEjer, @elab, @rev, @aut)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@no", r(2))
                    .Parameters.AddWithValue("@libro", r(3))
                    .Parameters.AddWithValue("@periodo", r(4))
                    .Parameters.AddWithValue("@anio", r(5))
                    .Parameters.AddWithValue("@ubicacion", r(6))
                    .Parameters.AddWithValue("@cveCorteEjer", r(7))
                    .Parameters.AddWithValue("@elab", r(8))
                    .Parameters.AddWithValue("@rev", r(9))
                    .Parameters.AddWithValue("@aut", r(10))
                End With

            Case "E.9"
                initialQuery = "INSERT INTO dbo.[E.9]
(Secretaria, Dirección, No, FechaAcuerdo, AcuerdoPend, UnidadAdmin, FechaReal, CveCorteEjer, Elaboró, Revisó, Autorizó) "
                initialQuery &= "VALUES
 (@sec, @dir, @no, @fechaAcuerdo, @acuerdoPend, @unidadAdmin, @fechaReal, @cveCorteEjer, @elaboro, @reviso, @autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@no", r(2))
                    .Parameters.AddWithValue("@fechaAcuerdo", r(3))
                    .Parameters.AddWithValue("@acuerdoPend", r(4))
                    .Parameters.AddWithValue("@unidadAdmin", r(5))
                    .Parameters.AddWithValue("@fechaReal", r(6))
                    .Parameters.AddWithValue("@cveCorteEjer", r(7))
                    .Parameters.AddWithValue("@elaboro", r(8))
                    .Parameters.AddWithValue("@reviso", r(9))
                    .Parameters.AddWithValue("@autorizo", r(10))
                End With

            Case "E.10"
                initialQuery = "INSERT INTO dbo.[E.10]
(Secretaria, Dirección, No, OrigenProg,NombProg,Periodo,TipoBenef,TotalBenef,Dependencia,CveCorteEjer, Elaboró, Revisó, Autorizó) "
                initialQuery &= "VALUES
 (@sec, @dir, @no,@OrigenProg,@NombProg,@Periodo,@TipoBenef,@TotalBenef,@Dependencia  ,@cveCorteEjer, @elaboro, @reviso, @autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@sec", r(0))
                    .Parameters.AddWithValue("@dir", r(1))
                    .Parameters.AddWithValue("@no", r(2))
                    .Parameters.AddWithValue("@OrigenProg", r(3))
                    .Parameters.AddWithValue("@NombProg", r(4))
                    .Parameters.AddWithValue("@Periodo", r(5))
                    .Parameters.AddWithValue("@TipoBenef", r(6))
                    .Parameters.AddWithValue("@TotalBenef", r(7))
                    .Parameters.AddWithValue("@Dependencia", r(8))
                    .Parameters.AddWithValue("@cveCorteEjer", r(9))
                    .Parameters.AddWithValue("@elaboro", r(10))
                    .Parameters.AddWithValue("@reviso", r(11))
                    .Parameters.AddWithValue("@autorizo", r(12))
                End With

            Case "F.1"
                initialQuery = "INSERT INTO dbo.[F.1]
(Secretaria, Dirección,  NombreExp,Período,TipoClasif,Ubicación,CveCorteEjer,Elaboró, Revisó, Autorizó) "
                initialQuery &= "VALUES
 (@Secretaria,@Dirección,@NombreExp,@Período,@TipoClasif,@Ubicación,@CveCorteEjer,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@NombreExp", r(2))
                    .Parameters.AddWithValue("@Período", r(3))
                    .Parameters.AddWithValue("@TipoClasif", r(4))
                    .Parameters.AddWithValue("@Ubicación", r(5))
                    .Parameters.AddWithValue("@CveCorteEjer", r(6))
                    .Parameters.AddWithValue("@Elaboró", r(7))
                    .Parameters.AddWithValue("@Revisó", r(8))
                    .Parameters.AddWithValue("@Autorizó", r(9))
                End With

            Case "F.1.1"
                initialQuery = "INSERT INTO dbo.[F.1.1]
(Secretaria, Dirección,  Código,NombreExp,Período,Ubicación,CveCorteEjer,Elaboró, Revisó, Autorizó) "
                initialQuery &= "VALUES
 (@Secretaria,@Dirección,@Código,@NombreExp,@Período,@Ubicación,@CveCorteEjer,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@Código", r(2))
                    .Parameters.AddWithValue("@NombreExp", r(3))
                    .Parameters.AddWithValue("@Período", r(4))
                    .Parameters.AddWithValue("@Ubicación", r(5))
                    .Parameters.AddWithValue("@CveCorteEjer", r(6))
                    .Parameters.AddWithValue("@Elaboró", r(7))
                    .Parameters.AddWithValue("@Revisó", r(8))
                    .Parameters.AddWithValue("@Autorizó", r(9))
                End With

            Case "F.2"
                initialQuery = "INSERT INTO dbo.[F.2]
(Secretaria, Dirección, Descipción,Fecha,Ubicación,CveCorteEjer, Elaboró, Revisó, Autorizó) "
                initialQuery &= "VALUES
 (@Secretaria,@Dirección,@Descipción,@Fecha,@Ubicación,@CveCorteEjer,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@Descipción", r(2))
                    .Parameters.AddWithValue("@Fecha", r(3))
                    .Parameters.AddWithValue("@Ubicación", r(4))
                    .Parameters.AddWithValue("@CveCorteEjer", r(5))
                    .Parameters.AddWithValue("@Elaboró", r(6))
                    .Parameters.AddWithValue("@Revisó", r(7))
                    .Parameters.AddWithValue("@Autorizó", r(8))
                End With

            Case "F.3"
                initialQuery = "INSERT INTO dbo.[F.3]
(Secretaria, Dirección,  Nombre,Fecha,Justificación,Observaciones,CveCorteEjer,Elaboró, Revisó, Autorizó) "
                initialQuery &= "VALUES
 (@Secretaria,@Dirección,@Nombre,@Fecha,@Justificación,@Observaciones,@CveCorteEjer,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@Nombre", r(2))
                    .Parameters.AddWithValue("@Fecha", r(3))
                    .Parameters.AddWithValue("@Justificación", r(4))
                    .Parameters.AddWithValue("@Observaciones", r(5))
                    .Parameters.AddWithValue("@CveCorteEjer", r(6))
                    .Parameters.AddWithValue("@Elaboró", r(7))
                    .Parameters.AddWithValue("@Revisó", r(8))
                    .Parameters.AddWithValue("@Autorizó", r(9))
                End With

            Case "F.4"
                initialQuery = "INSERT INTO dbo.[F.4]
(Secretaria, Dirección,  Descripción,Sello,CveCorteEjer,Elaboró, Revisó, Autorizó) "
                initialQuery &= "VALUES
 (@Secretaria,@Dirección,@Descripción,@Sello,@CveCorteEjer,@Elaboró,@Revisó,@Autorizó)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@Secretaria", r(0))
                    .Parameters.AddWithValue("@Dirección", r(1))
                    .Parameters.AddWithValue("@Descripción", r(2))
                    .Parameters.AddWithValue("@Sello", r(3))
                    .Parameters.AddWithValue("@CveCorteEjer", r(4))
                    .Parameters.AddWithValue("@Elaboró", r(5))
                    .Parameters.AddWithValue("@Revisó", r(6))
                    .Parameters.AddWithValue("@Autorizó", r(7))
                End With

            Case "I"
                initialQuery &= "INSERT INTO dbo.[I] (Secretaria, Dirección,No,InfoActiv,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@No,@InfoActiv,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("No", r(2))
                    .Parameters.AddWithValue("InfoActiv", r(3))
                    .Parameters.AddWithValue("CveCorteEjer", r(4))
                    .Parameters.AddWithValue("Elaboro", r(5))
                    .Parameters.AddWithValue("reviso", r(6))
                    .Parameters.AddWithValue("autorizo", r(7))
                End With

            Case "II"
                initialQuery &= "INSERT INTO dbo.[II] (Secretaria, Dirección,IdOranigrama,Organigrama,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@IdOranigrama,@Organigrama,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("IdOranigrama", r(2))
                    .Parameters.AddWithValue("Organigrama", r(3))
                    .Parameters.AddWithValue("CveCorteEjer", r(4))
                    .Parameters.AddWithValue("Elaboro", r(5))
                    .Parameters.AddWithValue("reviso", r(6))
                    .Parameters.AddWithValue("autorizo", r(7))
                End With

            Case "III"
                initialQuery &= "INSERT INTO dbo.[III] (Secretaria, Dirección,No,FunciónGral,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@No,@FunciónGral,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("No", r(2))
                    .Parameters.AddWithValue("FunciónGral", r(3))
                    .Parameters.AddWithValue("CveCorteEjer", r(4))
                    .Parameters.AddWithValue("Elaboro", r(5))
                    .Parameters.AddWithValue("reviso", r(6))
                    .Parameters.AddWithValue("autorizo", r(7))
                End With

            Case "IV"
                initialQuery &= "INSERT INTO dbo.[IV] (Secretaria, Dirección,SecciónAnexo,ClaveAnexo,NombreAnexo,Art28_LGMENL,Aplica,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@SecciónAnexo,@ClaveAnexo,@NombreAnexo,@Art28_LGMENL,@Aplica,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("SecciónAnexo", r(2))
                    .Parameters.AddWithValue("ClaveAnexo", r(3))
                    .Parameters.AddWithValue("NombreAnexo", r(4))
                    .Parameters.AddWithValue("Art28_LGMENL", r(5))
                    .Parameters.AddWithValue("Aplica", r(6))
                    .Parameters.AddWithValue("CveCorteEjer", r(7))
                    .Parameters.AddWithValue("Elaboro", r(8))
                    .Parameters.AddWithValue("reviso", r(9))
                    .Parameters.AddWithValue("autorizo", r(10))
                End With

            Case "V"
                initialQuery &= "INSERT INTO dbo.[V] (Secretaria, Dirección,No,PMD,CveCorteEjer, Elaboró, Revisó, Autorizó)"
                initialQuery &= "VALUES (@secretaria,@direccion,@No,@PMD,@CveCorteEjer,@Elaboro,@reviso,@autorizo)"
                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("@secretaria", r(0))
                    .Parameters.AddWithValue("direccion", r(1))
                    .Parameters.AddWithValue("No", r(2))
                    .Parameters.AddWithValue("PMD", r(3))
                    .Parameters.AddWithValue("CveCorteEjer", r(4))
                    .Parameters.AddWithValue("Elaboro", r(5))
                    .Parameters.AddWithValue("reviso", r(6))
                    .Parameters.AddWithValue("autorizo", r(7))
                End With

            Case "ActaER"
                initialQuery &= "INSERT INTO dbo.[ActaER] (CveCorteEjer,IdMunicipio,HoraInicioActa,MesInicioActa,DiaInicioActa,AñoInicioActa,AñoInicioActaLetra,ReunionEn,ColoniaReunion,CalleReunion,PteMunSaliente,CallePMS,NumeroExternoPMS,ColoniaPMS,RfcPMS,FolioInePMS,PteMunEntrante,CallePME,NumeroExternoPME,ColoniaPME,RfcPME,FolioInePME,PeriodoInicial,PeriodoFinal,TestigoPMS,CalleTestigoPMS,NumeroExternoTestigoPMS,ColoniaTestigoPMS,RfcTestigoPMS,FolioIneTestigoPMS,TestigoPME,CalleTestigoPME,NumeroExternoTestigoPME,ColoniaTestigoPME,RfcTestigoPME,FolioIneTestigoPME,Contralor,FolioINEContralor,Sindico1S,FolioIneSindico1S,Sindico1E,FolioIneSindico1E,Sindico2S,FolioIneSindico2S,Sindico2E,FolioIneSindico2E,NumeroAnexos,Al_DiaEntrega,De_MesEntrega,Año_Entrega,HoraFinActa,DiaFinActa,MesFinActa,AñoFinActa,AñoFinActaLetra)"
                initialQuery &= "VALUES (@CveCorteEjer,@IdMunicipio,@HoraInicioActa,@MesInicioActa,@DiaInicioActa,@AñoInicioActa,@AñoInicioActaLetra,@ReunionEn,@ColoniaReunion,@CalleReunion,@PteMunSaliente,@CallePMS,@NumeroExternoPMS,@ColoniaPMS,@NumeroExternoPMS,@ColoniaPMS,@RfcPMS,@FolioInePMS,@PteMunEntrante,@CallePME,@NumeroExternoPME,@ColoniaPME,@RfcPME,@FolioInePME,@PeriodoInicial,@PeriodoFinal,@TestigoPMS,@CalleTestigoPMS,@NumeroExternoTestigoPMS,@ColoniaTestigoPMS,@RfcTestigoPMS,@FolioIneTestigoPMS,@TestigoPME,@CalleTestigoPME,@NumeroExternoTestigoPME,@ColoniaTestigoPME,@RfcTestigoPME,@FolioIneTestigoPME,@Contralor,@FolioINEContralor,@Sindico1S,@FolioIneSindico1S,@Sindico1E,@FolioIneSindico1E,@Sindico2S,@FolioIneSindico2S,@Sindico2E,@FolioIneSindico2E,@NumeroAnexos,@Al_DiaEntrega,@De_MesEntrega,@Año_Entrega,@HoraFinActa,@DiaFinActa,@MesFinActa,@AñoFinActa,@AñoFinActaLetra)"

                With sql
                    .CommandText = initialQuery
                    .Parameters.AddWithValue("CveCorteEjer", r(1))
                    .Parameters.AddWithValue("IdMunicipio", r(2))
                    .Parameters.AddWithValue("HoraInicioActa", r(3))
                    .Parameters.AddWithValue("MesInicioActa", r(4))
                    .Parameters.AddWithValue("DiaInicioActa", r(5))
                    .Parameters.AddWithValue("AñoInicioActa", r(6))
                    .Parameters.AddWithValue("AñoInicioActaLetra", r(7))
                    .Parameters.AddWithValue("ReunionEn", r(8))
                    .Parameters.AddWithValue("ColoniaReunion", r(9))
                    .Parameters.AddWithValue("CalleReunion", r(10))
                    .Parameters.AddWithValue("PteMunSaliente", r(11))
                    .Parameters.AddWithValue("CallePMS", r(12))
                    .Parameters.AddWithValue("NumeroExternoPMS", r(13))
                    .Parameters.AddWithValue("ColoniaPMS", r(14))
                    .Parameters.AddWithValue("RfcPMS", r(15))
                    .Parameters.AddWithValue("FolioInePMS", r(16))
                    .Parameters.AddWithValue("PteMunEntrante", r(17))
                    .Parameters.AddWithValue("CallePME", r(18))
                    .Parameters.AddWithValue("NumeroExternoPME", r(19))
                    .Parameters.AddWithValue("ColoniaPME", r(20))
                    .Parameters.AddWithValue("RfcPME", r(21))
                    .Parameters.AddWithValue("FolioInePME", r(22))
                    .Parameters.AddWithValue("PeriodoInicial", r(23))
                    .Parameters.AddWithValue("PeriodoFinal", r(24))
                    .Parameters.AddWithValue("TestigoPMS", r(25))
                    .Parameters.AddWithValue("CalleTestigoPMS", r(26))
                    .Parameters.AddWithValue("NumeroExternoTestigoPMS", r(27))
                    .Parameters.AddWithValue("ColoniaTestigoPMS", r(28))
                    .Parameters.AddWithValue("RfcTestigoPMS", r(29))
                    .Parameters.AddWithValue("FolioIneTestigoPMS", r(30))
                    .Parameters.AddWithValue("TestigoPME", r(31))
                    .Parameters.AddWithValue("CalleTestigoPME", r(32))
                    .Parameters.AddWithValue("NumeroExternoTestigoPME", r(33))
                    .Parameters.AddWithValue("ColoniaTestigoPME", r(34))
                    .Parameters.AddWithValue("RfcTestigoPME", r(35))
                    .Parameters.AddWithValue("FolioIneTestigoPME", r(36))
                    .Parameters.AddWithValue("Contralor", r(37))
                    .Parameters.AddWithValue("FolioINEContralor", r(38))
                    .Parameters.AddWithValue("Sindico1S", r(39))
                    .Parameters.AddWithValue("FolioIneSindico1S", r(40))
                    .Parameters.AddWithValue("Sindico1E", r(41))
                    .Parameters.AddWithValue("FolioIneSindico1E", r(42))
                    .Parameters.AddWithValue("Sindico2S", r(43))
                    .Parameters.AddWithValue("FolioIneSindico2S", r(44))
                    .Parameters.AddWithValue("Sindico2E", r(45))
                    .Parameters.AddWithValue("FolioIneSindico2E", r(46))
                    .Parameters.AddWithValue("NumeroAnexos", r(47))
                    .Parameters.AddWithValue("Al_DiaEntrega", r(48))
                    .Parameters.AddWithValue("De_MesEntrega", r(49))
                    .Parameters.AddWithValue("Año_Entrega", r(50))
                    .Parameters.AddWithValue("HoraFinActa", r(51))
                    .Parameters.AddWithValue("DiaFinActa", r(52))
                    .Parameters.AddWithValue("MesFinActa", r(53))
                    .Parameters.AddWithValue("AñoFinActa", r(54))
                    .Parameters.AddWithValue("AñoFinActaLetra", r(55))

                End With





        End Select
    End Function

End Class
