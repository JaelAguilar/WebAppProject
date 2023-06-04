﻿Imports System.Data.SqlClient
Imports MigraDoc.DocumentObjectModel
Imports MigraDoc.DocumentObjectModel.Tables
Imports TheArtOfDev.HtmlRenderer.Core

Partial Public Class WebForm1
    Public Function CreatePDF() As Document


        Dim databaseName = exportTableSelector.SelectedValue

        Dim doc As New Document()
        doc.Info.Title = "Reporte"
        'Defining styles



        Dim styles As Style = doc.Styles.Item(StyleNames.Normal)
        styles.Font.Name = "Segoe UI"


        styles = doc.Styles("Heading1")
        styles.Font.Name = "Helvetica"
        styles.Font.Size = 18
        styles.Font.Bold = True
        styles.Font.Color = Colors.White
        styles.ParagraphFormat.SpaceBefore = 20
        styles.ParagraphFormat.Alignment = ParagraphAlignment.Center

        styles = doc.Styles("Heading2")
        styles.Font.Name = "Segoe UI"
        styles.Font.Size = 12
        styles.Font.Color = Colors.Black
        styles.ParagraphFormat.SpaceBefore = 15
        styles.ParagraphFormat.Alignment = ParagraphAlignment.Center


        doc.Add(CreateSection("A.1"))


        Return doc
    End Function

    Private Function CreateSection(databaseName As String)
        Dim secretaryName = exportSecretarySelector.SelectedValue
        Dim directoryName = exportDirectorySelector.SelectedValue

        'obtain Data from the database
        Dim conn As New SqlConnection
        conn.Open()
        Dim command As New SqlCommand()
        With command
            .CommandText = "select * from @table where secretaría=@secretaria and direccion=@direccion"
            .Parameters.AddWithValue("@table", databaseName)
            .Parameters.AddWithValue("@secretaria", secretaryName)
            .Parameters.AddWithValue("@direccion", directoryName)
        End With
        Dim reader As SqlDataReader = command.ExecuteReader
        Dim dt As New DataTable
        dt.Load(reader)
        Dim paragraph As Paragraph
        Dim page As New Section


        page.PageSetup.LeftMargin = Unit.FromInch(0.5)
        page.PageSetup.RightMargin = Unit.FromInch(0.5)
        page.PageSetup.TopMargin = Unit.FromInch(0.5)

        Dim tRow As Row
        Dim currentTable As Table




        ''Aquí voy a crear un ejemplo de cómo se verán los datos una vez se recuperen de la tabla de sql, esto es porque ustedes no tienen la base de datos, y se me hizo más sencillo que tuvieran ustedes mismos la tabla.
        ''Aquí pondrán los nombres de cada una de sus columnas de su tabla
        'Dim columns As New List(Of String)({"Secretaria", "Direccion", "I_PreEst", "I_PreMod", "I_PreDev", "I_PreRec", "E_PreOrigApro", "E_1A_AmPres", "E_2A_AmpPres", "E_3A_AmpPres", "E_Tot_Amp", "E_PreModif", "E_PreComp", "E_PreDev", "E_PreEjer", "E_PreErog", "E_PreCons", "E_PrePorEjer", "FechaCorte", "Elaboró", "Revisó", "Autorizó"})
        'Dim dbTest As New DataTable()

        'For index = 0 To columns.Count - 1
        '    dbTest.Columns.Add(columns(index))
        'Next
        ''Aquí pondrán sus datos de prueba para asegurarse que los datos correctos se muestran en el reporte,no importa mucho cuáles sean son sólo de prueba. Cada instrucción de dbTestRows.Add() Agrega una nueva fila
        'dbTest.Rows.Add("s1", "d1", "ipe1", "ipm1", "ipd1", "ipr1", "epoa1", "e1mp1", "e2mp1", "e3mp1", "eta1", "epm1", "epc1", "epd1", "epej1", "eper1", "epco1", "eppe1", "fc1", "e1", "r1", "a1")
        'dbTest.Rows.Add("s2", "d2", "ipe2", "ipm2", "ipd2", "ipr2", "epoa2", "e1mp2", "e2mp2", "e3mp2", "eta2", "epm2", "epc2", "epd2", "epej2", "eper2", "epco2", "eppe2", "fc2", "e2", "r2", "a2")
        Dim direccion = "D1"
        Dim secretaria = "S1"
        page.AddParagraph("Secretaría " & secretaria & ", Dirección " & direccion)

        Select Case databaseName
            Case "A.1"
                'Headings
                paragraph = page.AddParagraph("PRESUPUESTO GLOBAL 2023", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray


                page.AddParagraph("SE ANEXA INFORMACIÓN" & Environment.NewLine & "PRESUPUESTO DE INGRESOS Y EGRESOS GLOBALES" & Environment.NewLine & "(MILES DE PESOS)" & Environment.NewLine, "Heading2")

                page.AddParagraph(" ")

                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(0.4))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.Columns(1).Format.Alignment = ParagraphAlignment.Center 'The $ symbol is in the center


                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("INGRESOS")
                tRow(0).MergeRight = 2
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                '2nd Headings row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CONCEPTO")
                tRow(1).AddParagraph("IMPORTE")
                tRow(1).MergeRight = 1
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO ESTIMADO")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(2))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO MODIFICADO")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(3))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO DEVENGADO")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(4))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO RECAUDADO")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))


                page.Add(currentTable)
                'Empty space
                page.AddParagraph(" " & Environment.NewLine & " ")

                'Create second table
                currentTable = New Table()
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(0.4))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.Columns(1).Format.Alignment = ParagraphAlignment.Center

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("EGRESOS")
                tRow(0).MergeRight = 2
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center


                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CONCEPTO")
                tRow(1).AddParagraph("IMPORTE")
                tRow(1).MergeRight = 1
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center


                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO ORIGINAL APROBADO")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(6))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("1ER. AMPLIACIÓN PRESUPUESTAL (SE ANEXA DOC.)")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(7))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("2DA. AMPLIACIÓN PRESUPUESTAL (SE ANEXA DOC.)")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(8))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("3ER. AMPLIACIÓN PRESUPUESTAL (SE ANEXA DOC.)")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(9))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TOTAL AMPLIACIONES (2 + 3 + 4)")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(10))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO MODIFICADO (SE ANEXA DOC.)")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(11))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO COMPROMETIDO (SE ANEXA DOC.)")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(12))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO DEVENGADO (SE ANEXA DOC.)")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(13))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO EJERCIDO (SE ANEXA DOC.)")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO EROGADO (SE ANEXA DOC.)")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(15))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO CONSUMIDO (SE ANEXA DOC.) (7+8+9+10)")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(16))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO POR EJERCER OFICIAL (SE ANEXA DOC.) (6 -11)")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(17))
                page.Add(currentTable)


            Case "A.1.1"
                'Headings
                paragraph = page.AddParagraph("PRESUPUESTO GLOBAL 2023", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph("SE ANEXA INFORMACIÓN" & Environment.NewLine & "PRESUPUESTO DE INGRESOS Y EGRESOS GLOBALES" & Environment.NewLine & "(MILES DE PESOS)" & Environment.NewLine, "Heading2")
                page.AddParagraph(" ")
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                'Creating columns
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(0.4))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.Columns(1).Format.Alignment =
        ParagraphAlignment.Center 'The $ symbol is in the center
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("INGRESOS")
                tRow(0).MergeRight = 2
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                '2nd Headings row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CONCEPTO")
                tRow(1).AddParagraph("IMPORTE")
                tRow(1).MergeRight = 1
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                'Data
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("IMPUESTOS")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(2))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CUOTAS Y APORTACIONES DE SEGURIDAD SOCIAL")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(3))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CONTRIBUCIONES DE MEJORAS")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(4))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("DERECHOS")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRODUCTOS")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow
                tRow(0).AddParagraph("APROVECHAMIENTOS")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow
                tRow(0).AddParagraph("INGRESOS PARA VENTAS DE BIENES Y SERVICIOS")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow
                tRow(0).AddParagraph("PARTICIPACIONES Y APORTACIONES")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow
                tRow(0).AddParagraph("TRANSFERENCIAS, ASIGNACIONES, SUBSIDIOS Y OTRAS AYUDAS")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow
                tRow(0).AddParagraph("INGRESOS DERIVADOS DE FINANCIAMIENTO")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                page.Add(currentTable)
                'Empty space
                page.AddParagraph(" " & Environment.NewLine & " ")

            Case "A.1.2"
                paragraph = page.AddParagraph("PRESUPUESTO GLOBAL 2023", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph("SE ANEXA INFORMACIÓN" & Environment.NewLine & "PRESUPUESTO DE EGRESOS AUTORIZADO GLOBAL POR CAPITULOS" & Environment.NewLine & "(MILES DE PESOS)" & Environment.NewLine, "Heading2")
                page.AddParagraph(" ")
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                'Creating columns
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(0.4))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.Columns(1).Format.Alignment =
        ParagraphAlignment.Center 'The $ symbol is in the center
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("INGRESOS")
                tRow(0).MergeRight = 2
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                '2nd Headings row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CONCEPTO")
                tRow(1).AddParagraph("IMPORTE")
                tRow(1).MergeRight = 1
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                'Data
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("SERVICIOS PERSONALES")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(2))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("MATERIALES Y SUMINISTROS")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(3))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TRANSFERENCIAS, ASIGNACIONES, SUBSIDIOS Y OTRAS AYUDAS")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(4))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("BIENES MUEBLES, INMUEBLES E INTANGIBLES")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("INVERSION PÚBLICA")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow
                tRow(0).AddParagraph("INVERSIONES FINANCIERAS Y OTRAS PROVISIONES")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow
                tRow(0).AddParagraph("INGRESOS PARA VENTAS DE BIENES Y SERVICIOS")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow
                tRow(0).AddParagraph("PARTICIPACIONES Y APORTACIONES")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow
                tRow(0).AddParagraph("DEUDA PÚBLICA")
                tRow(1).AddParagraph("$")
                tRow(2).AddParagraph(dt(0)(5))
                page.Add(currentTable)

            Case "A.1.3"
                paragraph = page.AddParagraph("PRESUPUESTO GLOBAL 2023", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray

                page.AddParagraph("SE ANEXA INFORMACIÓN" & Environment.NewLine & "PRESUPUESTO DE EGRESOS AUTORIZADO GLOBAL POR CAPITULOS" & Environment.NewLine & "(MILES DE PESOS)" & Environment.NewLine, "Heading2")
                page.AddParagraph(" ")

                ' Create first table
                currentTable = New Table()

                ' Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                ' Creating columns
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1))

                ' Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("Clave")
                tRow(1).AddParagraph("Nombre")
                tRow(2).AddParagraph("Presupuesto Autorizado")
                tRow(3).AddParagraph("%")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                ' Data rows
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("301")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(6))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("302")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(7))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("303")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(8))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("304")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(9))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("305")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(10))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("306")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(11))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("307")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(12))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("308")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(13))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("309")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(15))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("310")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(16))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("311")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(17))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("312")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("313")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("314")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("315")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("316")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("317")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("318")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("319")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("320")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("321")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("322")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("323")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph(dt(0)(3))
                tRow(3).AddParagraph(dt(0)(14))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph(dt(0)(17))
                tRow(2).AddParagraph(dt(0)(0))
                tRow(3).AddParagraph(dt(0)(0))



                page.Add(currentTable)

                page.AddParagraph(" " & Environment.NewLine & " ")
                page.AddParagraph(" " & Environment.NewLine & " ")
                page.AddParagraph(" " & Environment.NewLine & " ")


                currentTable = New Table()
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                currentTable.AddColumn(Unit.FromInch(0.75))
                currentTable.AddColumn(Unit.FromInch(2.5))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.Columns(1).Format.Alignment =
ParagraphAlignment.Center

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("Total de gasto corriente")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(0).MergeRight = 1
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("113")
                tRow(1).AddParagraph("OBRA PÚBLICA DIRECTA")
                tRow(2).AddParagraph(dt(0)(6))
                tRow(3).AddParagraph(dt(0)(7))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("114")
                tRow(1).AddParagraph("ACTIVO FIJO")
                tRow(2).AddParagraph(dt(0)(6))
                tRow(3).AddParagraph(dt(0)(7))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("115")
                tRow(1).AddParagraph("AMORTIZACIÓN DE LA DEUDA")
                tRow(2).AddParagraph(dt(0)(6))
                tRow(3).AddParagraph(dt(0)(7))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("116")
                tRow(1).AddParagraph("RAMO 33")
                tRow(2).AddParagraph(dt(0)(6))
                tRow(3).AddParagraph(dt(0)(7))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("117")
                tRow(1).AddParagraph("OTROS EGRESOS")
                tRow(2).AddParagraph(dt(0)(6))
                tRow(3).AddParagraph(dt(0)(7))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TOTAL PRESUPUESTO AUTORIZADO")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(0).MergeRight = 1
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                page.Add(currentTable)

                ' Empty space
                page.AddParagraph(" " & Environment.NewLine & " ")

            Case "A.1.4"
                paragraph = page.AddParagraph("PRESUPUESTO DE EGRESOS POR SECRETARÍA",
"Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" " & Environment.NewLine & " ")
                page.AddParagraph("SE ANEXA INFORMACIÓN" &
                Environment.NewLine & "PRESUPUESTO DE EGRESOS POR SECRETARIA",
               "Heading2")
                page.AddParagraph(" ")
                page.AddParagraph("(MILES DE PESOS)", "Heading3")
                page.AddParagraph(" ")
                page.AddParagraph(" " & Environment.NewLine & " ")
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2.5))
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(2))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CONCEPTO:")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("CLAVE:")
                tRow(3).AddParagraph("")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Left
                page.Add(currentTable)
                page.AddParagraph(" " & Environment.NewLine & " ")
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 1
                currentTable.TopPadding = 6
                currentTable.BottomPadding = 6
                currentTable.LeftPadding = 4
                currentTable.AddColumn(Unit.FromInch(3.7))
                currentTable.AddColumn(Unit.FromInch(3.7))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO DE EGRESOS 2023")
                tRow(0).MergeRight = 1
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CONCEPTO")
                tRow(1).AddParagraph("IMPORTE")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO ORIGINAL APROBADO")
                tRow(1).AddParagraph(dt(0)(2))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("1ER AMPLIACIÓN PRESUPUESTAL (SE ANEXA DOC)")
                tRow(1).AddParagraph(dt(0)(3))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("2DA AMPLIACIÓN PRESUPUESTAL (SE ANEXA DOC)")
                tRow(1).AddParagraph(dt(0)(4))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("3RA AMPLIACIÓN PRESUPUESTAL (SE ANEXA DOC)")
                tRow(1).AddParagraph(dt(0)(5))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TOTAL AMPLIACIÓN (2+3+4)")
                tRow(1).AddParagraph(dt(0)(6))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO MODIFICADO (SE ANEXA DOC)")
                tRow(1).AddParagraph(dt(0)(7))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO COMPROMETIDO (SE ANEXA DOC)")
                tRow(1).AddParagraph(dt(0)(8))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO DEVENGADO (SE ANEXA DOC)")
                tRow(1).AddParagraph(dt(0)(9))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO EJERCIDO (SE ANEXA DOC)")
                tRow(1).AddParagraph(dt(0)(10))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO EROGADO (SE ANEXA DOC)")
                tRow(1).AddParagraph(dt(0)(11))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO CONSUMIDO (SE ANEXA DOC)(7+8+9+10)")
                tRow(1).AddParagraph(dt(0)(12))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("PRESUPUESTO POR EJERCER (SE ANEXA DOC) (6-11)")
                tRow(1).AddParagraph(dt(0)(13))
                page.Add(currentTable)

            Case "A.4"
                'Headings
                paragraph = page.AddParagraph("ASIGNACIÓN DEL FONDO ÚNICO DE OPERACIÓN", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" ")
                'Create first table
                currentTable = New Table()
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                'Creating columns
                currentTable.AddColumn(Unit.FromInch(2.45))
                currentTable.AddColumn(Unit.FromInch(2.45))
                currentTable.AddColumn(Unit.FromInch(2.45))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("SECRETARIA")
                tRow(1).AddParagraph("TITULAR")
                tRow(2).AddParagraph("MONTO")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                    tRow.Cells(2).AddParagraph(dt(index)(4))
                Next
                page.Add(currentTable)

            Case "A.4.1"
                'Agrega aquí tu código
                paragraph = page.AddParagraph("ARQUEO DEL FONDO ÚNICO DE OPERACIONES",
               "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" " & Environment.NewLine & " ")
                page.AddParagraph(" " & Environment.NewLine & " ")
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.5
                currentTable.TopPadding = 2
                currentTable.BottomPadding = 2
                currentTable.LeftPadding = 2
                currentTable.AddColumn(Unit.FromInch(3.7))
                currentTable.AddColumn(Unit.FromInch(3.7))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("FECHA")
                tRow(1).AddParagraph(dt(0)(2))
                tRow.Format.Alignment = ParagraphAlignment.Left
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("RESPONSABLE DEL FONDO")
                tRow(1).AddParagraph(dt(0)(3))
                tRow.Format.Alignment = ParagraphAlignment.Left
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("ENCARGADO DEL FONDO")
                tRow(1).AddParagraph(dt(0)(4))
                tRow.Format.Alignment = ParagraphAlignment.Left
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("MONTO AUTORIZADO")
                tRow(1).AddParagraph(dt(0)(5))
                tRow(0).Format.Alignment = ParagraphAlignment.Center
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TOTAL EFECTIVO")
                tRow(1).AddParagraph(dt(0)(6))
                tRow(0).Format.Alignment = ParagraphAlignment.Center
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TOTAL DE DOCUMENTOS")
                tRow(1).AddParagraph(dt(0)(7))
                tRow(0).Format.Alignment = ParagraphAlignment.Center
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TOTAL DE ARQUEO")
                tRow(1).AddParagraph(dt(0)(8))
                tRow(0).Format.Alignment = ParagraphAlignment.Center
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("VARIACIÓN")
                tRow(1).AddParagraph(dt(0)(9))
                tRow(0).Format.Alignment = ParagraphAlignment.Center
                tRow(0).Format.Font.Bold = True
                page.Add(currentTable)
                page.AddParagraph(" " & Environment.NewLine & " ")
                page.AddParagraph(" " & Environment.NewLine & " ")
                'TABLA 2
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.5
                currentTable.TopPadding = 2
                currentTable.BottomPadding = 2
                currentTable.LeftPadding = 2
                currentTable.AddColumn(Unit.FromInch(1.3))
                currentTable.AddColumn(Unit.FromInch(1.1))
                currentTable.AddColumn(Unit.FromInch(1.1))
                currentTable.AddColumn(Unit.FromInch(0.4)).Borders.Bottom.Color = Colors.White
                currentTable.AddColumn(Unit.FromInch(1.3))
                currentTable.AddColumn(Unit.FromInch(1.1))
                currentTable.AddColumn(Unit.FromInch(1.1))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("BILLETE")
                tRow(4).AddParagraph("MONEDA")
                tRow(0).MergeRight = 2
                tRow(4).MergeRight = 2
                tRow.Cells(3).Borders.Bottom.Color = Colors.White
                tRow.Cells(3).Borders.Top.Color = Colors.White
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("DENOMINACIÓN")
                tRow(1).AddParagraph("BILLETES")
                tRow(2).AddParagraph("CANTIDAD")
                tRow(3).AddParagraph("")
                tRow.Cells(3).Borders.Bottom.Color = Colors.White
                tRow.Cells(3).Borders.Top.Color = Colors.White
                tRow(4).AddParagraph("DENOMINACIÓN")
                tRow(5).AddParagraph("MONEDAS")
                tRow(6).AddParagraph("CANTIDAD")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                'FILA 1
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("$1000")
                tRow(1).AddParagraph(dt(0)(10))
                tRow(2).AddParagraph(dt(0)(10))
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("$20")
                tRow(5).AddParagraph(dt(0)(15))
                tRow(6).AddParagraph(dt(0)(15))
                'FILA 2
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("$500")
                tRow(1).AddParagraph(dt(0)(11))
                tRow(2).AddParagraph(dt(0)(11))
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("$10")
                tRow(5).AddParagraph(dt(0)(16))
                tRow(6).AddParagraph(dt(0)(16))
                'FILA 3
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("$200")
                tRow(1).AddParagraph(dt(0)(12))
                tRow(2).AddParagraph(dt(0)(12))
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("$5")
                tRow(5).AddParagraph(dt(0)(17))
                tRow(6).AddParagraph(dt(0)(17))
                'FILA 4
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("$100")
                tRow(1).AddParagraph(dt(0)(13))
                tRow(2).AddParagraph(dt(0)(13))
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("$2")
                tRow(5).AddParagraph(dt(0)(18))
                tRow(6).AddParagraph(dt(0)(18))
                'FILA 5
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("$50")
                tRow(1).AddParagraph(dt(0)(14))
                tRow(2).AddParagraph(dt(0)(14))
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("$1")
                tRow(5).AddParagraph(dt(0)(19))
                tRow(6).AddParagraph(dt(0)(19))
                'FILA 6
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("$50C")
                tRow(5).AddParagraph(dt(0)(20))
                tRow(6).AddParagraph(dt(0)(20))
                'FILA 7
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("$20C")
                tRow(5).AddParagraph(dt(0)(21))
                tRow(6).AddParagraph(dt(0)(21))
                'FILA 8
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("$10C")
                tRow(5).AddParagraph(dt(0)(22))
                tRow(6).AddParagraph(dt(0)(22))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("SUBTOTAL")
                tRow(0).Format.Font.Bold = True
                tRow(0).MergeRight = 2
                tRow(4).MergeRight = 2
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("SUBTOTAL")
                tRow(4).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TOTAL")
                tRow(0).Format.Font.Bold = True
                tRow(0).MergeRight = 6
                page.Add(currentTable)
                page.AddParagraph(" " & Environment.NewLine & " ")
                page.AddParagraph(" " & Environment.NewLine & " ")
                'TABLA 3
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 3
                currentTable.BottomPadding = 3
                currentTable.LeftPadding = 3
                currentTable.AddColumn(Unit.FromInch(1.8))
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1.4))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NUMERO DE FACTURAS Y/O REMISIONES")
                tRow(1).AddParagraph("FECHA")
                tRow(2).AddParagraph("PROVEEDOR")
                tRow(3).AddParagraph("CONCEPTO")
                tRow(4).AddParagraph("IMPORTE")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("")
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("")
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("")
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("")
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TOTAL")
                tRow(0).Format.Font.Bold = True
                tRow(0).MergeRight = 4
                page.Add(currentTable)
                page.AddParagraph(" " & Environment.NewLine & " ")
                page.AddParagraph(" " & Environment.NewLine & " ")
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 8
                currentTable.BottomPadding = 8
                currentTable.LeftPadding = 8
                currentTable.AddColumn(Unit.FromInch(7.4))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("OBSERVACIONES")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Left
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph(dt(0)(23))
                page.Add(currentTable)



            Case "A.5"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE CUENTAS BANCARIAS, INVERSIONES, ETC.", "Heading1")
                Paragraph.Format.Borders.Width = 2.5
                Paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray

                page.AddParagraph(" ")

                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns
                currentTable.AddColumn(Unit.FromInch(3.8))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.5))



                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NOMBRE DE LA INSTITUCIÓN")
                tRow(1).AddParagraph("No DE CUENTA O CONTRATO")
                tRow(2).AddParagraph("TIPO - INVERSIÓN")
                tRow(3).AddParagraph("SALDO")
                tRow(4).AddParagraph("FECHA DE VENCIMIENTO")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center


                'Data
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                    tRow.Cells(2).AddParagraph(dt(index)(4))
                    tRow.Cells(3).AddParagraph(dt(index)(5))
                    tRow.Cells(4).AddParagraph(dt(index)(6))
                Next
                page.Add(currentTable)

            Case "A.5.1"
                paragraph = page.AddParagraph("DETALLE DE LAS CUENTAS DE CHEQUE", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" " & Environment.NewLine & " ")
                page.AddParagraph("SE ANEXA INFORMACIÓN", "Heading2")
                page.AddParagraph(" ")
                page.AddParagraph(" " & Environment.NewLine & " ")
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                currentTable.AddColumn(Unit.FromInch(3.7))
                currentTable.AddColumn(Unit.FromInch(0.925))
                currentTable.AddColumn(Unit.FromInch(0.925))
                currentTable.AddColumn(Unit.FromInch(0.925))
                currentTable.AddColumn(Unit.FromInch(0.925))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NOMBRE DE LA INSTITUCIÓN")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(1).MergeRight = 3
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NUMERO DE LA CUENTA DE CHEQUES")
                tRow(1).AddParagraph(dt(0)(3))
                tRow(1).MergeRight = 3
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CUENTA CONTABLE")
                tRow(1).AddParagraph(dt(0)(4))
                tRow(1).MergeRight = 3
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("SALDO SEGÚN LIBROS")
                tRow(1).AddParagraph(dt(0)(5))
                tRow(1).MergeRight = 3
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("SALDO SEGÚN ESTADOS DE CUENTA BANCARIO")
                tRow(1).AddParagraph(dt(0)(6))
                tRow(1).MergeRight = 3
                tRow(0).Format.Font.Bold = True
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("SALDO SEGÚN ESTADOS DE CUENTA BANCARIO")
                tRow(1).AddParagraph("DEL NO.")
                tRow(2).AddParagraph(dt(0)(7))
                tRow(3).AddParagraph("AL NO.")
                tRow(4).AddParagraph(dt(0)(8))
                tRow(0).Format.Font.Bold = True
                tRow(1).Format.Font.Bold = True
                tRow(3).Format.Font.Bold = True
                page.Add(currentTable)
                page.AddParagraph(" " & Environment.NewLine & " ")
                page.AddParagraph("RELACIÓN DE ÚLTIMOS (5) CHEQUES EXPEDIDOS:",
               "Heading2")
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1.2))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("FECHA")
                tRow(1).AddParagraph("NO.DE CHEQUE")
                tRow(2).AddParagraph("BENEFICIARIO")
                tRow(3).AddParagraph("CONCEPTO")
                tRow(4).AddParagraph("IMPORTE")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("")
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("")
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("")
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("")
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow(1).AddParagraph("")
                tRow(2).AddParagraph("")
                tRow(3).AddParagraph("")
                tRow(4).AddParagraph("")
                page.Add(currentTable)
                page.AddParagraph(" " & Environment.NewLine & " ")
                page.AddParagraph(" " & Environment.NewLine & " ")
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.Borders.Bottom.Color = Colors.White
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                currentTable.AddColumn(Unit.FromInch(3.7))
                currentTable.AddColumn(Unit.FromInch(3.7))
                page.AddParagraph("FIRMAS REGISTRADAS", "Heading2")
                page.AddParagraph(" " & Environment.NewLine & " ")
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("FIRMA")
                tRow.Cells(0).Borders.Top.Color = Colors.White
                tRow.Cells(0).Borders.Bottom.Color = Colors.White
                tRow.Cells(0).Borders.Left.Color = Colors.White
                tRow.Cells(0).Borders.Right.Color = Colors.White
                tRow(1).AddParagraph("FIRMA")
                tRow.Cells(1).Borders.Top.Color = Colors.White
                tRow.Cells(1).Borders.Bottom.Color = Colors.White
                tRow.Cells(1).Borders.Right.Color = Colors.White
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph(dt(0)(9))
                tRow(0).AddParagraph("______________")
                tRow.Cells(0).Borders.Bottom.Color = Colors.White
                tRow(1).AddParagraph(dt(0)(10))
                tRow(1).AddParagraph("______________")
                tRow.Cells(1).Borders.Bottom.Color = Colors.White
                tRow.Format.Font.Bold = True
                tRow.Cells(0).Borders.Left.Color = Colors.White
                tRow.Cells(0).Borders.Right.Color = Colors.White
                tRow.Cells(1).Borders.Right.Color = Colors.White
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NOMBRE")
                tRow.Cells(0).Borders.Bottom.Color = Colors.White
                tRow(1).AddParagraph("NOMBRE")
                tRow.Cells(1).Borders.Bottom.Color = Colors.White
                tRow.Format.Font.Bold = True
                tRow.Cells(0).Borders.Left.Color = Colors.White
                tRow.Cells(0).Borders.Right.Color = Colors.White
                tRow.Cells(1).Borders.Right.Color = Colors.White
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CARGO")
                tRow.Cells(0).Borders.Bottom.Color = Colors.White
                tRow(1).AddParagraph("CARGO")
                tRow.Cells(1).Borders.Bottom.Color = Colors.White
                tRow.Cells(0).Borders.Left.Color = Colors.White
                tRow.Cells(0).Borders.Right.Color = Colors.White
                tRow.Cells(1).Borders.Right.Color = Colors.White
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("")
                tRow.Cells(0).Borders.Bottom.Color = Colors.White
                tRow(1).AddParagraph("")
                tRow.Cells(1).Borders.Bottom.Color = Colors.White
                tRow.Cells(0).Borders.Left.Color = Colors.White
                tRow.Cells(0).Borders.Right.Color = Colors.White
                tRow.Cells(1).Borders.Right.Color = Colors.White
                '2
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("FIRMA")
                tRow.Cells(0).Borders.Bottom.Color = Colors.White
                tRow(1).AddParagraph("FIRMA")
                tRow.Cells(1).Borders.Bottom.Color = Colors.White
                tRow.Format.Font.Bold = True
                tRow.Cells(0).Borders.Left.Color = Colors.White
                tRow.Cells(0).Borders.Right.Color = Colors.White
                tRow.Cells(1).Borders.Right.Color = Colors.White
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph(dt(0)(11))
                tRow(0).AddParagraph("______________")
                tRow.Cells(0).Borders.Bottom.Color = Colors.White
                tRow(1).AddParagraph(dt(0)(12))
                tRow(1).AddParagraph("______________")
                tRow.Cells(1).Borders.Bottom.Color = Colors.White
                tRow.Format.Font.Bold = True
                tRow.Cells(0).Borders.Left.Color = Colors.White
                tRow.Cells(0).Borders.Right.Color = Colors.White
                tRow.Cells(1).Borders.Right.Color = Colors.White
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NOMBRE")
                tRow.Cells(0).Borders.Bottom.Color = Colors.White
                tRow(1).AddParagraph("NOMBRE")
                tRow.Cells(1).Borders.Bottom.Color = Colors.White
                tRow.Format.Font.Bold = True
                tRow.Cells(0).Borders.Left.Color = Colors.White
                tRow.Cells(0).Borders.Right.Color = Colors.White
                tRow.Cells(1).Borders.Right.Color = Colors.White
                tRow.Format.Alignment = ParagraphAlignment.Center
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CARGO")
                tRow.Cells(0).Borders.Bottom.Color = Colors.White
                tRow(1).AddParagraph("CARGO")
                tRow.Cells(1).Borders.Bottom.Color = Colors.White
                tRow.Cells(0).Borders.Left.Color = Colors.White
                tRow.Cells(0).Borders.Right.Color = Colors.White
                tRow.Cells(1).Borders.Right.Color = Colors.White
                tRow.Format.Alignment = ParagraphAlignment.Center
                page.Add(currentTable)
                page.AddParagraph(" " & Environment.NewLine & " ")
                page.AddParagraph(" " & Environment.NewLine & " ")
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(1.66))
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(1.66))
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(1.66))
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph(dt(0)(13))
                tRow(1).AddParagraph("INDIVIDUAL")
                tRow(2).AddParagraph(dt(0)(14))
                tRow(3).AddParagraph("MANCOMUNADA")
                tRow(4).AddParagraph(dt(0)(15))
                tRow(5).AddParagraph("INDISTINTA")
                tRow.Format.Font.Bold = True
                tRow.Cells(1).Borders.Top.Color = Colors.White
                tRow.Cells(1).Borders.Bottom.Color = Colors.White
                tRow.Cells(3).Borders.Top.Color = Colors.White
                tRow.Cells(3).Borders.Bottom.Color = Colors.White
                tRow.Cells(5).Borders.Top.Color = Colors.White
                tRow.Cells(5).Borders.Bottom.Color = Colors.White
                tRow.Cells(5).Borders.Right.Color = Colors.White
                tRow.Format.Alignment = ParagraphAlignment.Center
                page.Add(currentTable)

            Case "A.6"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings
                paragraph = page.AddParagraph("RELACION DE CUENTAS POR COBRAR", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray

                page.AddParagraph(" ")

                'Create first table
                currentTable = New Table()

                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(4.8))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. DE DOCUMENTO")
                tRow(1).AddParagraph("NOMBRE DEL DEUDOR")
                tRow(2).AddParagraph("FECHA DE ADEUDO")
                tRow(3).AddParagraph("SALDO")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                    tRow.Cells(2).AddParagraph(dt(index)(4))
                    tRow.Cells(3).AddParagraph(dt(index)(5))
                Next

                page.Add(currentTable)

            Case "A.6.2"
                paragraph = page.AddParagraph("REPORTE DE SALDOS DEL SISTEMA DE INGRESOS", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray

                'Create first table
                currentTable = New Table()

                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(4))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NÚMERO")
                tRow(1).AddParagraph("TIPO DE CUENTA")
                tRow(2).AddParagraph("SALDO A LA FECHA")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                    tRow.Cells(2).AddParagraph(dt(index)(4))
                Next

                page.Add(currentTable)

            Case "A.7"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings
                paragraph = page.AddParagraph("RELACION DE CUENTAS POR PAGAR", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray

                page.AddParagraph(" ")

                'Create first table
                currentTable = New Table()

                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns
                currentTable.AddColumn(Unit.FromInch(2.9))
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(2.9))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. DE DOCUMENTO")
                tRow(1).AddParagraph("NOMBRE DEL ACREEDOR")
                tRow(2).AddParagraph("SALDO")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                    tRow.Cells(2).AddParagraph(dt(index)(4))
                Next

                page.Add(currentTable)

            Case "A.7.1"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE SALDOS CON PROVEEDORES Y CONTRATISTAS", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" ")

                'Create first table
                currentTable = New Table()

                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns Suma 10.8
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.3))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("SECRETARÍA")
                tRow(1).AddParagraph("DIRECCIÓN")
                tRow(2).AddParagraph("PROVEEDOR")
                tRow(3).AddParagraph("SALDO")
                tRow(4).AddParagraph("CLAVE EJERCICIO")
                tRow(5).AddParagraph("ELABORÓ")
                tRow(6).AddParagraph("REVISÓ")
                tRow(7).AddParagraph("AUTORIZÓ")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                page.Add(currentTable)

            Case "A.7.2"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE SALDOS CON ACREEDORES DIVERSOS", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" ")

                'Create first table
                currentTable = New Table()

                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns Suma 10.8
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.3))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("SECRETARÍA")
                tRow(1).AddParagraph("DIRECCIÓN")
                tRow(2).AddParagraph("ACREEDOR")
                tRow(3).AddParagraph("SALDO")
                tRow(4).AddParagraph("CLAVE CORTE")
                tRow(5).AddParagraph("ELABORÓ")
                tRow(6).AddParagraph("REVISÓ")
                tRow(7).AddParagraph("AUTORIZÓ")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                page.Add(currentTable)

            Case "A.7.3"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE DOCUMENTOS POR PAGAR", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" ")

                'Create first table
                currentTable = New Table()

                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2.8))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. DE DOCUMENTO")
                tRow(1).AddParagraph("NOMBRE DEL ACREEDOR")
                tRow(2).AddParagraph("FECHA")
                tRow(3).AddParagraph("SALDO")
                tRow(4).AddParagraph("VENCIMIENTO")
                tRow(5).AddParagraph("CONCEPTO")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                page.Add(currentTable)

        End Select
        Return page

    End Function
End Class
