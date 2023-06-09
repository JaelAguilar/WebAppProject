﻿Imports System.Buffers.Text
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports MigraDoc.DocumentObjectModel
Imports MigraDoc.DocumentObjectModel.Shapes
Imports MigraDoc.DocumentObjectModel.Tables
Imports PdfSharp.Pdf.IO
Imports TheArtOfDev.HtmlRenderer.Core
Imports Orientation = MigraDoc.DocumentObjectModel.Orientation

Partial Public Class WebForm1
    Public Function CreatePDF() As Document


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


        Dim lItems = ListBoxReport.GetSelectedIndices()
        For Each it In lItems
            Dim name = ListBoxReport.Items(it).Value
            doc.Add(RetrieveOriginalPDF(name))
            doc.Add(CreateSection(name))
        Next

        Return doc
    End Function

    Private Function RetrieveData(databaseName As String) As DataTable
        Dim secretaryName = exportSecretary.Text
        Dim directoryName = exportDirectory.Text

        'obtain Data from the database
        Dim conn As New SqlConnection("server=WIN-2VFJL7TQ7Q8\SQLEXPRESS;database=WebDatabase;User ID='externaluser';Password='public12##'")
        Dim dt As New DataTable

        Try
            conn.Open()
            Dim command As New SqlCommand()
            With command
                .Connection = conn
                .CommandType = CommandType.Text
                .CommandText = "select * from dbo.[" & databaseName & "] where Secretaria=@secretaria and Dirección=@direccion"
                .Parameters.AddWithValue("@table", databaseName)
                .Parameters.AddWithValue("@secretaria", secretaryName)
                .Parameters.AddWithValue("@direccion", directoryName)
            End With
            Using reader2 = command.ExecuteReader
                dt.Load(reader2)
                Return dt
            End Using
        Catch ex As Exception
            'MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Function

    Private Function RetrieveOriginalPDF(databaseName As String)
        Dim inputDocument = PdfReader.Open(AppDomain.CurrentDomain.BaseDirectory & "reports\" & databaseName & ".pdf")
        Dim p = New Section
        Dim cover = p.AddImage(AppDomain.CurrentDomain.BaseDirectory & "reports\" & databaseName & ".pdf#1")
        cover.RelativeVertical = RelativeVertical.Page
        cover.RelativeHorizontal = RelativeHorizontal.Page
        cover.WrapFormat.Style = WrapStyle.Through

        Return p
    End Function

    Private Function CreateSection(databaseName As String)

        Dim dt = RetrieveData(databaseName)

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
        'Dim dt As New DataTable()

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

            Case "A.2"
                'Headings 
                paragraph = page.AddParagraph("REPORTE DE PRESUPUESTO OPERATIVO", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(2.4))

                'Heading row 
                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("CONCEPTO")
                tRow(1).AddParagraph("IMPORTE")
                tRow.Format.Font.Bold = True


                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("PRESUPUESTO ORIGINAL APROBADO")
                tRow(1).AddParagraph(dt(0)(2))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("AMPLIACIÓN Y/O REDUCCIONES")
                tRow(1).AddParagraph(dt(0)(3))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("PRESUPUESTO MODIFICADO (SE ANEXA DOC.)")
                tRow(1).AddParagraph(dt(0)(4))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("PRESUPUESTO COMPROMETIDO (SE ANEXA DOC.)")
                tRow(1).AddParagraph(dt(0)(5))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("PRESUPUESTO DEVENGADO (SE ANEXA DOC.)")
                tRow(1).AddParagraph(dt(0)(7))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("PRESUPUESTO EJERCIDO (SE ANEXA DOC)")
                tRow(1).AddParagraph(dt(0)(8))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("PRESUPUESTO EROGADO (SE ANEXA DOC.)")
                tRow(1).AddParagraph(dt(0)(9))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("PRESUPUESTO CONSUMIDO (SE ANEXA DOC. SOPORTE) (4 + 5 + 6 + 7)")
                tRow(1).AddParagraph(dt(0)(10))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("PRESUPUESTO POR EJERCER OFICIAL (3 –8)")
                tRow(1).AddParagraph(dt(0)(11))
                page.Add(currentTable)

            Case "A.3"
                'Headings 
                paragraph = page.AddParagraph("ESTADOS FINANCIEROS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(2.4))

                'Heading row 


                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("A. ESTADO DE SITUACIÓN FINANCIERA DE ACTIVIDADES")
                tRow(0).Format.Font.Bold = True
                tRow(1).AddParagraph(dt(0)(2))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("B. CUENTA PÚBLICA DEL EJERCICIO 2020")
                tRow(0).Format.Font.Bold = True
                tRow(1).AddParagraph(dt(0)(4))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("C. DICTAMEN DE AUDITORES EXTERNOS Y NOTAS AL ESTADO FINANCIERO REALIZADOS POR DESPACHO")
                tRow(0).Format.Font.Bold = True
                tRow(1).AddParagraph(dt(0)(5))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("ÚLTIMO PERÍODO DICTAMINADO:")
                tRow(0).Format.Font.Bold = True
                tRow(1).AddParagraph(dt(0)(6))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("D. BALANZA DE COMPROBACIÓN AL ")
                tRow(0).Format.Font.Bold = True
                tRow(1).AddParagraph(dt(0)(7))

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
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
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

            Case "A.5.1.1"


                'Create first table 
                currentTable = New Table()

                'Style 
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(2.4))

                'Heading row 
                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("CONCILIACÍON DE CUENTAS DE CHEQUES")
                tRow(1).AddParagraph("CLAVE DEL EJERCICIO")

                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data 
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                Next

                page.Add(currentTable)

            Case "A.5.2"
                'Headings 
                paragraph = page.AddParagraph("DETALLE DE LAS CUENTAS DE INVERSIÓN", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(2.4))

                'Heading row 
                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("NOMBRE DE LA INSTITUCIÓN")
                tRow(1).AddParagraph(dt(0)(2))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("CUENTA DE INVERSIONES")
                tRow(1).AddParagraph(dt(0)(3))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("CUENTA CONTABLE")
                tRow(1).AddParagraph(dt(0)(4))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("SALDO SEGÚN LIBROS")
                tRow(1).AddParagraph(dt(0)(5))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("SALDO SEGÚN ESTADOS DE CUENTA BANCARIO")
                tRow(1).AddParagraph(dt(0)(6))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("TIPOS DE INVERSIÓN")
                tRow(1).AddParagraph(dt(0)(7))

                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("VENCIMIENTO")
                tRow(1).AddParagraph(dt(0)(8))

                page.Add(currentTable)

            Case "A.5.2.1"
                'Headings 
                paragraph = page.AddParagraph("CONCILIACIÓN DE LAS CUENTAS DE INVERSIÓN", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(2.4))

                'Heading row 
                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("CONCILIACÍON DE CUENTAS DE INVERSIÓN")
                tRow(1).AddParagraph("CLAVE DEL EJERCICIO")

                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data 
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                Next

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



                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(0))
                    tRow.Cells(1).AddParagraph(dt(index)(1))
                    tRow.Cells(2).AddParagraph(dt(index)(2))
                    tRow.Cells(3).AddParagraph(dt(index)(3))
                    tRow.Cells(4).AddParagraph(dt(index)(4))
                    tRow.Cells(5).AddParagraph(dt(index)(5))
                    tRow.Cells(6).AddParagraph(dt(index)(6))
                    tRow.Cells(7).AddParagraph(dt(index)(7))
                Next
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



                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(0))
                    tRow.Cells(1).AddParagraph(dt(index)(1))
                    tRow.Cells(2).AddParagraph(dt(index)(2))
                    tRow.Cells(3).AddParagraph(dt(index)(3))
                    tRow.Cells(4).AddParagraph(dt(index)(4))
                    tRow.Cells(5).AddParagraph(dt(index)(5))
                    tRow.Cells(6).AddParagraph(dt(index)(6))
                    tRow.Cells(7).AddParagraph(dt(index)(7))
                Next
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

                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                    tRow.Cells(2).AddParagraph(dt(index)(4))
                    tRow.Cells(3).AddParagraph(dt(index)(5))
                    tRow.Cells(4).AddParagraph(dt(index)(6))
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                Next
                page.Add(currentTable)

            Case "A.8"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE CHEQUES PENDIENTES DE ENTREGAR.", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2.5))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2.5))
                currentTable.AddColumn(Unit.FromInch(1.2))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("FECHA")
                tRow(1).AddParagraph("No DE CUENTA BANCARIA")
                tRow(2).AddParagraph("INSTITUCION")
                tRow(3).AddParagraph("No DE CHEQUE")
                tRow(4).AddParagraph("NOMBRE DEL BENEFICIARIO")
                tRow(5).AddParagraph("IMPORTE")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                Next
                page.Add(currentTable)

            Case "A.9"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE  POLIZAS DE FIANZA QUE GARANTIZAN UN CREDITO FISCAL.", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1.3))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(2.5))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. de Poliza")
                tRow(1).AddParagraph("Nombre la compañia afianzadora")
                tRow(2).AddParagraph("Nombre de la compañia afianzadora")
                tRow(3).AddParagraph("Nombre del deudor")
                tRow(4).AddParagraph("Monto")
                tRow(5).AddParagraph("Concepto")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                Next
                page.Add(currentTable)

            Case "A.10"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("ESTADO DE GUARDA LA CUENTA PUBLICA.", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.8))
                currentTable.AddColumn(Unit.FromInch(1.5))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No.")
                tRow(1).AddParagraph("EJERCICIO FISCAL")
                tRow(2).AddParagraph("ACTA DE ENTREGA LA ASENL")
                tRow(3).AddParagraph("NUMERO DE LEGAJOS")
                tRow(4).AddParagraph("NUMERO DE DISCOS")
                tRow(5).AddParagraph("ESTATUS")
                tRow(6).AddParagraph("OBSERVACIONES O REQUERIMENTOS DE LA ASENL")
                tRow(6).AddParagraph("RESPONSABLE")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(9))
                Next
                page.Add(currentTable)

            Case "B.1"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE PERSONAL DE BASE, POR HONORARIOS ASIMILABLES A SUELDO Y/O TEMPORALES", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1.4))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NÚMERO DE NÓMINA")
                tRow(1).AddParagraph("PUESTO")
                tRow(2).AddParagraph("NOMBRE COMPLETO")
                tRow(3).AddParagraph("SUELDO Y/O PERCEPCIÓN")
                tRow(4).AddParagraph("VIGENCIA")
                tRow(5).AddParagraph("SINDICALIZADOS")
                tRow(6).AddParagraph("NO SINDICALIZADOS")
                tRow(7).AddParagraph("HONORARIOS ASIMILABLES")
                tRow(8).AddParagraph("TEMPORALES Y/O EVENTUALES")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(9))
                    tRow.Cells(8).AddParagraph(dt(index)(10))
                Next
                page.Add(currentTable)

            Case "B.2"

                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE PERSONAL DE BASE, POR HONORARIOS ASIMILABLES A SUELDO Y/O TEMPORALES", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(0.5))
                currentTable.AddColumn(Unit.FromInch(1.9))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NÚMERO DE NÓMINA")
                tRow(1).AddParagraph("SITUACIÓN")
                tRow(2).AddParagraph("LUGAR DE LA COMISIÓN")
                tRow(3).AddParagraph("CANTIDAD DE DÍAS Y/O PERÍODO")
                tRow(4).AddParagraph("OBSERVACIONES")
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

                page.AddParagraph("CLAVES DE SITUACIÓN" & Environment.NewLine & "L.LICENCIA    P.PERMISO    C.COMISIÓN    I.INCAPACIDAD    S.SUSPENCIÓN    V.VACACIONES" & Environment.NewLine, "Heading2")
                page.AddParagraph(" ")

            Case "B.3"
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE TURNOS Y CANTIDAD DE PERSONAS ASIGNADAS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(3.7))
                currentTable.AddColumn(Unit.FromInch(3.7))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TURNO")
                tRow(1).AddParagraph("NÚMERO DE EMPLEADOS")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                'Data
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                Next
                page.Add(currentTable)

            Case "B.4"
                'Agrega aquí tu código

                'ENCABEZADO
                paragraph = page.AddParagraph("RELACIÓN DE PERSONAL PENSIONADO Y JUBILADO", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" ")

                'TABLA
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns (7.4in PARA VERTICAL)
                currentTable.AddColumn(Unit.FromInch(1.6))
                currentTable.AddColumn(Unit.FromInch(2.4))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.4))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NÚMERO DE NÓMINA")
                tRow(1).AddParagraph("NOMBRE COMPLETO")
                tRow(2).AddParagraph("CLASIFICACIÓN")
                tRow(3).AddParagraph("PERCEPCIÓN MENSUAL")
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



                'PIE DE PAGINA
                page.AddParagraph("TOTAL DE PENSIONADOS:" &
                          Environment.NewLine & "TOTAL DE JUBILADOS:" &
                          Environment.NewLine, "Heading2")
                page.AddParagraph(" ")

            Case "C.1"
                'ORIENTACION 
                page.PageSetup.Orientation = Orientation.Landscape

                'ENCABEZADO
                paragraph = page.AddParagraph("RELACIÓN DE MOBILIARIO, EQUIPO, HERRAMIENTAS, ACCESORIOS Y/O LISTADO DE PATRIMONIO", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" ")

                'TABLA
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns (10.8in PARA HORIZONTAL)
                currentTable.AddColumn(Unit.FromInch(3))
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(2.8))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NÚMERO DE INVENTARIO")
                tRow(1).AddParagraph("DESCRIPCIÓN")
                tRow(2).AddParagraph("No NOMINA DEL RESGUARDANTE")
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



                'PIE DE PAGINA
                page.AddParagraph("TOTAL DE BIENES:" &
                                  Environment.NewLine, "Heading2")
                page.AddParagraph(" ")

            Case "C.2"
                'ORIENTACION 
                page.PageSetup.Orientation = Orientation.Landscape

                'ENCABEZADO
                paragraph = page.AddParagraph("RELACIÓN DE EQUIPO DE TRANSPORTE, MAQUINARIA Y COMBUSTIBLE ASIGNADO A VEHÍCULOS", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" ")

                'TABLA
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns (10.8in PARA HORIZONTAL)
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(0.7))
                currentTable.AddColumn(Unit.FromInch(0.7))
                currentTable.AddColumn(Unit.FromInch(0.7))
                currentTable.AddColumn(Unit.FromInch(0.6))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(0.9))
                currentTable.AddColumn(Unit.FromInch(0.7))
                currentTable.AddColumn(Unit.FromInch(0.7))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. INVENTARIO")
                tRow(1).AddParagraph("DESCRIPCIÓN")
                tRow(2).AddParagraph("MARCA")
                tRow(3).AddParagraph("MODELO")
                tRow(4).AddParagraph("No. PLACA")
                tRow(5).AddParagraph("No. SERIE")
                tRow(6).AddParagraph("No. NÓMINA RESGUARDANTE")
                tRow(7).AddParagraph("CONDICIONES")
                tRow(8).AddParagraph("TIPO Y CANT. COMBUSTIBLE")
                tRow(8).MergeRight = 1
                tRow(9).AddParagraph(" ")
                tRow(10).AddParagraph("ESTACIÓN ASIGNADA")
                tRow(11).AddParagraph("No. POLIZA Y COBERTURA")
                tRow(11).MergeRight = 1
                tRow(12).AddParagraph(" ")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(9))
                    tRow.Cells(8).AddParagraph(dt(index)(10))
                    tRow.Cells(9).AddParagraph(dt(index)(11))
                    tRow.Cells(10).AddParagraph(dt(index)(12))
                    tRow.Cells(11).AddParagraph(dt(index)(13))
                    tRow.Cells(12).AddParagraph(dt(index)(14))
                Next

                page.Add(currentTable)



                'PIE DE PAGINA
                page.AddParagraph("TOTAL DE BIENES:" &
                                  Environment.NewLine & "CLAVES PARA TIPO DE COBERTURA DE LA PÓLIZA: CA.  COBERTURA AMPLIA o RC.  RESPONSABILIDAD CIVIL" &
                                  Environment.NewLine, "Heading2")
                page.AddParagraph(" ")

            Case "C.3"
                'ENCABEZADO
                paragraph = page.AddParagraph("RELACIÓN DE LEYES, REGLAMENTOS, MANUALES, LIBROS Y PUBLICACIONES", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" ")

                'TABLA
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns (7.4in PARA VERTICAL)
                currentTable.AddColumn(Unit.FromInch(4))
                currentTable.AddColumn(Unit.FromInch(3.4))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TÍTULO Y/O CLASIFICACIÓN")
                tRow(1).AddParagraph("NÚMERO DE EJEMPLARES")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                Next

                page.Add(currentTable)



                'PIE DE PAGINA
                page.AddParagraph("TOTAL DE BIENES:" &
                                  Environment.NewLine, "Heading2")
                page.AddParagraph(" ")

            Case "C.4"
                'ENCABEZADO
                paragraph = page.AddParagraph("RELACIÓN DE PAPELERÍA OFICIAL EN STOCK", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" ")

                'TABLA
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns (7.4in PARA VERTICAL)
                currentTable.AddColumn(Unit.FromInch(1.1))
                currentTable.AddColumn(Unit.FromInch(1.7))
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(3))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CANTIDAD")
                tRow(1).AddParagraph("NÚMERO DE FORMATO")
                tRow(2).AddParagraph("NÚMERO DE FOLIO INICIAL / FINAL")
                tRow(2).MergeRight = 1
                tRow(3).AddParagraph(" ")
                tRow(4).AddParagraph("OBSERVACIONES")
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

            Case "C.5"
                'ORIENTACION 
                page.PageSetup.Orientation = Orientation.Landscape

                'ENCABEZADO
                paragraph = page.AddParagraph("INVENTARIO DE ALMACENES", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" ")

                'TABLA
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns (10.8in PARA HORIZONTAL)
                currentTable.AddColumn(Unit.FromInch(1.6))
                currentTable.AddColumn(Unit.FromInch(3.2))
                currentTable.AddColumn(Unit.FromInch(1.6))
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(2))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CÓDIGO")
                tRow(1).AddParagraph("DESCRIPCIÓN")
                tRow(2).AddParagraph("CANTIDAD")
                tRow(3).AddParagraph("CONDICIONES" + vbCrLf + "BUENO / REGULAR / MALO")
                tRow(3).MergeRight = 2
                tRow(4).AddParagraph(" ")
                tRow(5).AddParagraph(" ")
                tRow(6).AddParagraph("OBSERVACIONES")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                Next

                page.Add(currentTable)

            Case "C.1.5"
                'ORIENTACION 
                page.PageSetup.Orientation = Orientation.Landscape

                'ENCABEZADO
                paragraph = page.AddParagraph("RELACIÓN DE BIENES INTANGIBLES", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                page.AddParagraph(" ")

                'TABLA
                'Create first table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4

                'Creating columns (10.8in PARA HORIZONTAL)
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.1))
                currentTable.AddColumn(Unit.FromInch(1.1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.1))
                currentTable.AddColumn(Unit.FromInch(1.1))
                currentTable.AddColumn(Unit.FromInch(1.1))
                currentTable.AddColumn(Unit.FromInch(1.1))
                currentTable.AddColumn(Unit.FromInch(1.1))
                currentTable.AddColumn(Unit.FromInch(1.1))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. INVENTARIO")
                tRow(1).AddParagraph("NOMBRE DEL SOFTWARE")
                tRow(2).AddParagraph("PROVEEDOR")
                tRow(3).AddParagraph("No. SERIE")
                tRow(4).AddParagraph("No. LICENCIA")
                tRow(5).AddParagraph("DISCOS ORIGINALES")
                tRow(6).AddParagraph("VERSIÓN")
                tRow(7).AddParagraph("ÁREA RESPONSABLE")
                tRow(8).AddParagraph("No. RESGUARDANTE")
                tRow(9).AddParagraph("UBICACIÓN")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(9))
                    tRow.Cells(8).AddParagraph(dt(index)(10))
                    tRow.Cells(9).AddParagraph(dt(index)(11))
                Next

                page.Add(currentTable)



                'PIE DE PAGINA
                page.AddParagraph("TOTAL DE BIENES:" &
                                  Environment.NewLine, "Heading2")
                page.AddParagraph(" ")

            Case "C.2"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACION DE EQUIPO DE TRANSPORTE, MAQUINARIA Y COMBUSTIBLE ASIGNADO A VEHICULO.", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(3.8))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.5))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. DE INVENTARIO.")
                tRow(1).AddParagraph("DESCRIPCIÓN")
                tRow(2).AddParagraph("MARCA")
                tRow(3).AddParagraph("MODELO")
                tRow(4).AddParagraph("No. DE PLACA")
                tRow(5).AddParagraph("No. DE SERIE")
                tRow(6).AddParagraph("No. DE NÓMINA DEL RESGUARDANTE")
                tRow(7).AddParagraph("CONDICIONES")
                tRow(8).AddParagraph("TIPO Y CANT. DE COMBUSTIBLE")
                tRow(9).AddParagraph("ESTACIÓN ASIGNADA")
                tRow(10).AddParagraph("PÓLIZA")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(9))
                    tRow.Cells(8).AddParagraph(dt(index)(10))
                    tRow.Cells(9).AddParagraph(dt(index)(11))
                    tRow.Cells(10).AddParagraph(dt(index)(12))
                Next
                page.Add(currentTable)

            Case "C.3"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACION DE MANUALES DE ORGANIZACION Y DE PROCESOS.", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(3.8))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.5))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No.")
                tRow(1).AddParagraph("NOMBRE")
                tRow(2).AddParagraph("CODIGO Y REVISIÓN")
                tRow(3).AddParagraph("FECHA DE APROBACIÓN")
                tRow(4).AddParagraph("DEPENDENCIA QUE GENERÓ EL MANUAL")
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

            Case "C.3.1"
                paragraph = page.AddParagraph("RELACION DE LEYES, REGLAMENTOS, MANUALES, LIBROS Y PUBLICACIONES",
"Heading1")
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
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(0.4))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.Columns(1).Format.Alignment =
        ParagraphAlignment.Center

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TITULO Y/O CLASIFICACION")
                tRow(1).AddParagraph("NUMERO DE EJEMPLARES")

                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))

                Next
                page.Add(currentTable)

            Case "C.4"
                paragraph = page.AddParagraph("RELACION DE PAPELERIA OFICIAL EN STOCK", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                'Create table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2.2))
                currentTable.AddColumn(Unit.FromInch(0.5))
                currentTable.AddColumn(Unit.FromInch(0.5))
                currentTable.AddColumn(Unit.FromInch(2.2))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CANTIDAD")
                tRow(1).AddParagraph("NOMBRE DEL FORMATO")
                tRow(2).AddParagraph("INICIAL")
                tRow(3).AddParagraph("FINAL")
                tRow(4).AddParagraph("OBSERVACIONES")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                page.Add(currentTable)

                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                    tRow.Cells(2).AddParagraph(dt(index)(4))
                    tRow.Cells(3).AddParagraph(dt(index)(5))
                    tRow.Cells(4).AddParagraph(dt(index)(6))
                Next
                page.Add(currentTable)

            Case "C.5"
                paragraph = page.AddParagraph("INVENTARIO DE ALMACENES", "Heading1")
                paragraph.Format.Borders.Width = 2.5
                paragraph.Format.Borders.Color = Colors.Black
                paragraph.Format.Borders.Distance = 3
                paragraph.Format.Shading.Color = Colors.Gray
                'Create table
                currentTable = New Table()
                'Style
                currentTable.Borders.Width = 0.75
                currentTable.TopPadding = 4
                currentTable.BottomPadding = 4
                currentTable.LeftPadding = 4
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1.4))
                currentTable.AddColumn(Unit.FromInch(1.4))

                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CODIGO")
                tRow(1).AddParagraph("DESCRIPCION ")
                tRow(2).AddParagraph("CANTIDAD")
                tRow(3).AddParagraph("CONDICIONES")
                tRow(4).AddParagraph("OBSERVACIONES")
                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center
                page.Add(currentTable)
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                    tRow.Cells(2).AddParagraph(dt(index)(4))
                    tRow.Cells(3).AddParagraph(dt(index)(5))
                    tRow.Cells(4).AddParagraph(dt(index)(6))
                Next
                page.Add(currentTable)



            Case "C.6"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE BIENES MUEBLES PROPIEDAD DE TERCEROS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(3.8))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.5))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CANTIDAD")
                tRow(1).AddParagraph("DESCRIPCION BIEN MUEBLE")
                tRow(2).AddParagraph("NUMERO DE INVENTARIO O NUMERO DE SERIE")
                tRow(3).AddParagraph("NOMBRE DEL PROPIETARIO")
                tRow(4).AddParagraph("OBSERVACIONES")
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


            Case "C.7"
                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE ARMAMENTO MUNICIPAL Y DEL ESTADO", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1.4))

                'Heading row 
                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("NOMBRE DEL RESGUARDANTE")
                tRow(1).AddParagraph("TIPO DE ARMA")
                tRow(2).AddParagraph("CALIBRE")
                tRow(3).AddParagraph("NÚMERO DE SERIE O MATRÍCULA")
                tRow(4).AddParagraph("MUNICIPAL O ESTATAL")



                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data 
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(3))
                    tRow.Cells(1).AddParagraph(dt(index)(4))
                    tRow.Cells(2).AddParagraph(dt(index)(5))
                    tRow.Cells(3).AddParagraph(dt(index)(6))
                    tRow.Cells(4).AddParagraph(dt(index)(7))

                Next

                page.Add(currentTable)

            Case "C.8"
                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE CD'S Y CASETTES DE AUTIO Y VIDEO", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1.8))
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(3))
                currentTable.AddColumn(Unit.FromInch(1.8))

                'Heading row 
                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("TIPO DE FORMATO")
                tRow(1).AddParagraph("CANTIDAD")
                tRow(2).AddParagraph("NOMBRE DEL EVENTO")
                tRow(3).AddParagraph("FECHA DEL EVENTO")

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

            Case "C.9"
                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE LIBROS PROPIEDAD DEL GOBIERNO DEL ESTADO", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(4))
                currentTable.AddColumn(Unit.FromInch(3.4))

                'Heading row 
                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("CLASIFICACIÓN Y/O MATERIA")
                tRow(1).AddParagraph("NÚMERO DE EJEMPLARES")

                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data 
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                Next

                page.Add(currentTable)


            Case "C.10"
                'Agrega aquí tu código

                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE EQUINOS Y CANINOS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1.3))
                currentTable.AddColumn(Unit.FromInch(1.3))
                currentTable.AddColumn(Unit.FromInch(3))
                currentTable.AddColumn(Unit.FromInch(1.3))
                currentTable.AddColumn(Unit.FromInch(1.3))
                currentTable.AddColumn(Unit.FromInch(1.3))
                currentTable.AddColumn(Unit.FromInch(1.3))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. DE INVENTARIO")
                tRow(1).AddParagraph("EQUINO O CANINO")
                tRow(2).AddParagraph("NOMBRE")
                tRow(3).AddParagraph("FIERRO O CHIP")
                tRow(4).AddParagraph("DESCRIPCIÓN")
                tRow(5).AddParagraph("FECHA DE NACIMIENTO Y/O EDAD")
                tRow(6).AddParagraph("FECHA DE ADQUISICIÓN")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                Next
                page.Add(currentTable)

                page.AddParagraph("TOTAL DE EQUINOS:______________________ TOTAL DE CANINOS:______________________" & Environment.NewLine, "Heading2")
                page.AddParagraph(" ")

            Case "C.11"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE BIENES INMUEBLES PROPIEDAD DEL MUNICIPIO EN TRAMITE DE INCORPORACIÓN", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1.95))
                currentTable.AddColumn(Unit.FromInch(3))
                currentTable.AddColumn(Unit.FromInch(1.95))
                currentTable.AddColumn(Unit.FromInch(1.95))
                currentTable.AddColumn(Unit.FromInch(1.95))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. DE EXPEDIENTE")
                tRow(1).AddParagraph("USO")
                tRow(2).AddParagraph("UBICACIÓN")
                tRow(3).AddParagraph("SUPERFICIE")
                tRow(4).AddParagraph("ESTATUS")
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

            Case "C.11.1"
                'Orientación
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE BIENES INMUEBLES PROPIEDAD DEL MUNICIPIO ACREDITADOS Y/O INCORPORADOS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(2.8))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))


                'Heading row 

                tRow = currentTable.AddRow()


                tRow(0).AddParagraph("No. DE EXPEDIENTE")
                tRow(1).AddParagraph("USO")
                tRow(2).AddParagraph("UBICACIÓN")
                tRow(3).AddParagraph("SUPERFICIE")
                tRow(4).AddParagraph("DOCUMENTO QUE ACREDITA PROPIEDAD")


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

            Case "C.11.2"
                'Orientación
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE BIENES INMUEBLES EN COMODATO", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(2.8))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))


                'Heading row 

                tRow = currentTable.AddRow()


                tRow(0).AddParagraph("No. DE EXPEDIENTE")
                tRow(1).AddParagraph("USO")
                tRow(2).AddParagraph("UBICACIÓN")
                tRow(3).AddParagraph("NOMBRE DEL COMODATARIO")
                tRow(4).AddParagraph("VIGENCIA")


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

            Case "D.1"
                'Headings 
                paragraph = page.AddParagraph("PADRÓN DE PROVEEDORES Y CONTRATISTAS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(3.5))
                currentTable.AddColumn(Unit.FromInch(3.9))

                'Heading row 

                tRow = currentTable.AddRow()


                tRow(0).AddParagraph("CONTRATISTA")
                tRow(1).AddParagraph("ESPECIALIDAD")


                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data 
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))

                Next

                page.Add(currentTable)

            Case "D.2"
                'Orientación
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE OBRAS TERMINADAS Y EN PROCESO", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.8))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))

                'Heading row 

                tRow = currentTable.AddRow()


                tRow(0).AddParagraph("No. DE CONTRATO")
                tRow(1).AddParagraph("DESCRIPCIÓN DE LA OBRA")
                tRow(2).AddParagraph("CONTRATISTA ASIGNADO")
                tRow(3).AddParagraph("MONTO DE LA OBRA")
                tRow(4).AddParagraph("MONTO DE LO EJERCIDO")
                tRow(5).AddParagraph("MODALIDAD EN QUE SE CONTRATÓ LA OBRA")
                tRow(6).AddParagraph("UBICACIÓN DEL EXPEDIENTE")
                tRow(7).AddParagraph("RECURSOS UTILIZADOS")
                tRow(8).AddParagraph("% AVANCE")
                tRow(9).AddParagraph("AMP. DE METAS")


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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(9))
                    tRow.Cells(8).AddParagraph(dt(index)(10))
                    tRow.Cells(9).AddParagraph(dt(index)(11))

                Next

                page.Add(currentTable)

            Case "D.3"
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE PROGRAMAS Y PROYECTOS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1.6))
                currentTable.AddColumn(Unit.FromInch(2.2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.6))


                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. DE INDICADOR POA´S Y PBR")
                tRow(1).AddParagraph("NOMBRE DEL PROGRAMA O PROYECTO")
                tRow(2).AddParagraph("DESCRIPCIÓN")
                tRow(3).AddParagraph("PORCENTAJE DE CUMPLIMIENTO")
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


            Case "D.4"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE CONTRATOS  FINANCIADOS CON RECURSOS ESTATALES", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.75))
                currentTable.AddColumn(Unit.FromInch(1.75))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.75))
                currentTable.AddColumn(Unit.FromInch(1.25))
                currentTable.AddColumn(Unit.FromInch(1.5))

                'Heading row 
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No.")
                tRow(1).AddParagraph("No.DE CONTRATO")
                tRow(2).AddParagraph("CONTRATACIÓN (ADQUISIONES, OBRAS, SERVICIOS, OTROS")
                tRow(3).AddParagraph("MODALIDAD DE CONTRATACIÓN")
                tRow(4).AddParagraph("MONTO")
                tRow(5).AddParagraph("FUENTES DE FINANCIAMIENTO")
                tRow(6).AddParagraph("FECHA DE CONCLUSIÓN")
                tRow(7).AddParagraph("OBSERVACIONES")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(9))

                Next
                page.Add(currentTable)
            Case "D.5"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE CONTRATOS  FINANCIADOS CON RECURSOS FEDERALES", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.75))
                currentTable.AddColumn(Unit.FromInch(1.75))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.75))
                currentTable.AddColumn(Unit.FromInch(1.25))
                currentTable.AddColumn(Unit.FromInch(1.5))

                'Heading row 
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No.")
                tRow(1).AddParagraph("No.DE CONTRATO")
                tRow(2).AddParagraph("CONTRATACIÓN (ADQUISIONES, OBRAS, SERVICIOS, OTROS")
                tRow(3).AddParagraph("MODALIDAD DE CONTRATACIÓN")
                tRow(4).AddParagraph("MONTO")
                tRow(5).AddParagraph("FUENTES DE FINANCIAMIENTO")
                tRow(6).AddParagraph("FECHA DE CONCLUSIÓN")
                tRow(7).AddParagraph("OBSERVACIONES")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(9))

                Next
                page.Add(currentTable)

            Case "D.6"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE CONTRATOS  FINANCIADOS CON RECURSOS PROPIOS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(0.8))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.75))
                currentTable.AddColumn(Unit.FromInch(1.75))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.75))
                currentTable.AddColumn(Unit.FromInch(1.25))
                currentTable.AddColumn(Unit.FromInch(1.5))

                'Heading row 
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No.")
                tRow(1).AddParagraph("No.DE CONTRATO")
                tRow(2).AddParagraph("CONTRATACIÓN (ADQUISIONES, OBRAS, SERVICIOS, OTROS")
                tRow(3).AddParagraph("MODALIDAD DE CONTRATACIÓN")
                tRow(4).AddParagraph("MONTO")
                tRow(5).AddParagraph("FUENTES DE FINANCIAMIENTO")
                tRow(6).AddParagraph("FECHA DE CONCLUSIÓN")
                tRow(7).AddParagraph("OBSERVACIONES")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(9))

                Next
                page.Add(currentTable)

            Case "D.7"

                page.PageSetup.Orientation = Orientation.Landscape

                'Headings
                paragraph = page.AddParagraph("EXPEDIENTES DE OBRAS Y UBICACIÓN", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(4))
                currentTable.AddColumn(Unit.FromInch(2.2))
                currentTable.AddColumn(Unit.FromInch(2.2))
                currentTable.AddColumn(Unit.FromInch(2.4))


                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. DE CONTRATO")
                tRow(1).AddParagraph("VIGENCIA")
                tRow(2).AddParagraph("UBICACIÓN DE LA OBRA")
                tRow(3).AddParagraph("UBICACIÓN DEL EXPEDIENTE")
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

            Case "D.8"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE COMITÉS DE OBRA PÚBLICA FORMADOS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(4))
                currentTable.AddColumn(Unit.FromInch(2.2))
                currentTable.AddColumn(Unit.FromInch(2.2))
                currentTable.AddColumn(Unit.FromInch(2.4))


                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No.")
                tRow(1).AddParagraph("FECHA DE FORMACIÓN")
                tRow(2).AddParagraph("NÚMERO DE INTEGRANTES")
                tRow(3).AddParagraph("NÚMERO DE CONTRATO QUE CORRESPONDE")
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



            Case "E.1"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE AMPAROS , JUICIOS CONTENCIOSOS, ASUNTOS PENALES, ASUNTOS CIVILES Y LABORALES", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.6))
                currentTable.AddColumn(Unit.FromInch(2))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NÚMERO DE EXPEDIENTE")
                tRow(1).AddParagraph("No. DE JUZGADO O AGENCIA")
                tRow(2).AddParagraph("QUEJOSO, PROMOVENTE Y/O DEMANDANTE")
                tRow(3).AddParagraph("AUTORIDAD RESPONSABLE")
                tRow(4).AddParagraph("DEMANDADO")
                tRow(5).AddParagraph("ESTADO PROCESAL")
                tRow(6).AddParagraph("CONCEPTO DE LA DEMANDA Y/O IMPUGNACIÓN")
                tRow(7).AddParagraph("OBSERVACIONES")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(9))
                Next

                page.Add(currentTable)

            Case "E.2"
                'Headings
                paragraph = page.AddParagraph("ACUERDOS, CONTRATOS Y CONVENIOS VIGENTES", "Heading1")
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
                currentTable.Borders.Visible = False
                'Creating columns
                currentTable.AddColumn(Unit.FromInch(1.6))
                currentTable.AddColumn(Unit.FromInch(2.1))
                currentTable.AddColumn(Unit.FromInch(1.6))
                currentTable.AddColumn(Unit.FromInch(2.1))

                '1 row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TIPO: ")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph("FECHA  DE SUSCRIPCIÓN: ")
                tRow(3).AddParagraph(dt(0)(3))
                tRow(0).Format.Font.Bold = True
                tRow(2).Format.Font.Bold = True
                '2 row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("DURACIÓN: ")
                tRow(1).AddParagraph(dt(0)(4))
                tRow(2).AddParagraph("ENTIDAD: ")
                tRow(3).AddParagraph(dt(0)(5))
                tRow(0).Format.Font.Bold = True
                tRow(2).Format.Font.Bold = True
                '3 row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("DESCRIPCIÓN: ")
                tRow(0).MergeRight = 3
                tRow(0).Format.Font.Bold = True
                '4 row respuesta
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph(dt(0)(6))
                tRow(0).MergeRight = 3
                '5 row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("OBJETO: ")
                tRow(0).MergeRight = 3
                tRow(0).Format.Font.Bold = True
                '6 row respuesta
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph(dt(0)(7))
                tRow(0).MergeRight = 3
                '7 row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("SITUACIÓN: ")
                tRow(0).MergeRight = 3
                tRow(0).Format.Font.Bold = True
                '8 row respuesta
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph(dt(0)(8))
                tRow(0).MergeRight = 3

                page.Add(currentTable)

            Case "E.3"
                'Headings
                paragraph = page.AddParagraph("CONSEJOS, COMITÉS, FIDEICOMISOS, PATRONATOS, ASOCIACIONES Y HERMANAMIENTOS", "Heading1")
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
                currentTable.Borders.Visible = False
                'Creating columns
                currentTable.AddColumn(Unit.FromInch(1.3))
                currentTable.AddColumn(Unit.FromInch(1.7))
                currentTable.AddColumn(Unit.FromInch(1.9))
                currentTable.AddColumn(Unit.FromInch(2.5))

                '1 row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("TIPO: ")
                tRow(1).AddParagraph(dt(0)(2))
                tRow(2).AddParagraph("PERSONAS QUE LO FORMAN: ")
                tRow(3).AddParagraph(dt(0)(3))
                tRow(0).Format.Font.Bold = True
                tRow(2).Format.Font.Bold = True
                '2 row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("DURACIÓN: ")
                tRow(1).AddParagraph(dt(0)(4))
                tRow(2).AddParagraph("FECHA DE SUSCRIPCIÓN: ")
                tRow(3).AddParagraph(dt(0)(5))
                tRow(0).Format.Font.Bold = True
                tRow(2).Format.Font.Bold = True
                '3 row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("DESCRIPCIÓN: ")
                tRow(0).MergeRight = 3
                tRow(0).Format.Font.Bold = True
                '4 row respuesta
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph(dt(0)(6))
                tRow(0).MergeRight = 3
                '5 row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("OBJETO: ")
                tRow(0).MergeRight = 3
                tRow(0).Format.Font.Bold = True
                '6 row respuesta
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph(dt(0)(7))
                tRow(0).MergeRight = 3
                '7 row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("SITUACIÓN: ")
                tRow(0).MergeRight = 3
                tRow(0).Format.Font.Bold = True
                '8 row respuesta
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph(dt(0)(8))
                tRow(0).MergeRight = 3

                page.Add(currentTable)

            Case "E.4"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE DELEGADOS MUNICIPALES", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(3.8))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.5))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("SECCIÓN Y/O CLAVE")
                tRow(1).AddParagraph("NOMBRE")
                tRow(2).AddParagraph("DOMICILIO")
                tRow(3).AddParagraph("FECHA DE NOMBRAMIENTO")
                tRow(4).AddParagraph("OBSERVACIONES")
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


            Case "E.5"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE BIENES EMBARGADOS Y DECOMISADOS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(3.8))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.5))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NOMBRE DEL CONTRIBUYENTE")
                tRow(1).AddParagraph("CANTIDAD")
                tRow(2).AddParagraph("DESCRIPCIÓN DEL BIEN")
                tRow(3).AddParagraph("CLASIFICACIÓN")
                tRow(4).AddParagraph("MOTIVO")
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

            Case "E.6"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE INMUEBLES DESAFECTADOS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1.3))
                currentTable.AddColumn(Unit.FromInch(2))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. DE EXPEDIENTE")
                tRow(1).AddParagraph("NOMBRE DE POSTOR")
                tRow(2).AddParagraph("SUPERFICIE DESAFECTADA")
                tRow(3).AddParagraph("AUTORIZO CABILDO")
                tRow(4).AddParagraph("AUTORIZO CONGRESO")
                tRow(5).AddParagraph("TIPO ENAJENACIÓN")
                tRow(6).AddParagraph("No. DE DECRETO Y FECHA")
                tRow(7).AddParagraph("OBSERVACIONES")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(9))
                Next
                page.Add(currentTable)

            Case "E.7"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE REGULARIZACIÓN DE COLONIAS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(2.3))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(0.5))
                currentTable.AddColumn(Unit.FromInch(0.5))
                currentTable.AddColumn(Unit.FromInch(2))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No. DE EXPEDIENTE")
                tRow(1).AddParagraph("COLONIA Y/O UNIÓN")
                tRow(2).AddParagraph("No. DE DECRETO")
                tRow(3).AddParagraph("LOTES DESAFECTADOS")
                tRow(4).AddParagraph("E")
                tRow(5).AddParagraph("S/E")
                tRow(6).AddParagraph("OBSERVACIONES")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))

                Next
                page.Add(currentTable)

            Case "E.8"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE ACTAS DE CABILDO Y UBICACIÓN", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(0.5))
                currentTable.AddColumn(Unit.FromInch(3.6))
                currentTable.AddColumn(Unit.FromInch(3.7))
                currentTable.AddColumn(Unit.FromInch(1))
                currentTable.AddColumn(Unit.FromInch(2))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No.")
                tRow(1).AddParagraph("LIBRO DE ACTAS")
                tRow(2).AddParagraph("PERIODO DE LA ADMINISTRACIÓN MUNICIPAL")
                tRow(3).AddParagraph("AÑO")
                tRow(4).AddParagraph("UBICACIÓN")
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


            Case "E.9"
                page.PageSetup.Orientation = Orientation.Landscape
                'Headings
                paragraph = page.AddParagraph("ASUNTOS EN TRÁMITE DE LAS COMISIONES DE CABILDO", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(3.8))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(2))
                currentTable.AddColumn(Unit.FromInch(1.5))
                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No.")
                tRow(1).AddParagraph("FECHA EN QUE SE SIGNO EL ACUERDO")
                tRow(2).AddParagraph("ACUERDO PENDIENTE")
                tRow(3).AddParagraph("UNIDAD ADMINISTRATIVA RESPONSABLE")
                tRow(4).AddParagraph("FECHA O PLAZO A REALIZAR")
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


            Case "E.10"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings
                paragraph = page.AddParagraph("BENEFICIADOS DE LOS PROGRAMAS FEDERALES Y ESTATALES", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.2))
                currentTable.AddColumn(Unit.FromInch(1.6))
                currentTable.AddColumn(Unit.FromInch(2))

                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("No")
                tRow(1).AddParagraph("ORIGEN DEL PROGRAMA")
                tRow(2).AddParagraph("NOMBRE DEL PROGRAMA")
                tRow(3).AddParagraph("PERIODO")
                tRow(4).AddParagraph("TIPO DE BENEFICIADO")
                tRow(5).AddParagraph("TOTAL DE BENEFICIADOS")
                tRow(6).AddParagraph("DEPENDENCIA")
                tRow(7).AddParagraph("SECRETARIA")
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
                    tRow.Cells(5).AddParagraph(dt(index)(7))
                    tRow.Cells(6).AddParagraph(dt(index)(8))
                    tRow.Cells(7).AddParagraph(dt(index)(0))
                Next
                page.Add(currentTable)


            Case "F.1"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE PAPELERÍA OFICIAL EN USO Y EN ARCHIVO MUERTO", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(4))
                currentTable.AddColumn(Unit.FromInch(2.2))
                currentTable.AddColumn(Unit.FromInch(2.2))
                currentTable.AddColumn(Unit.FromInch(2.4))


                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("NOMBRE DEL EXPEDIENTE")
                tRow(1).AddParagraph("PERÍODO")
                tRow(2).AddParagraph("TIPO DE CLASIFICACIÓN")
                tRow(3).AddParagraph("UBICACIÓN")
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


            Case "F.1.1"
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings
                paragraph = page.AddParagraph("RELACIÓN DE EXPEDIENTES Y ACTAS EN ARCHIVO", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(4))
                currentTable.AddColumn(Unit.FromInch(2.2))
                currentTable.AddColumn(Unit.FromInch(2.2))
                currentTable.AddColumn(Unit.FromInch(2.4))


                'Heading row
                tRow = currentTable.AddRow()
                tRow(0).AddParagraph("CODIGO")
                tRow(1).AddParagraph("NOMBRE DEL EXPEDIENTE O ACTA")
                tRow(2).AddParagraph("PERIODO")
                tRow(3).AddParagraph("UBICACIÓN")
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

            Case "F.2"
                'Headings 
                paragraph = page.AddParagraph("ARCHIVO DE PLANOS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(4.4))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(1.5))

                'Heading row 
                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("DESCRIPCIÓN DEL PLANO")
                tRow(1).AddParagraph("FECHA")
                tRow(2).AddParagraph("UBICACIÓN")

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

            Case "F.3"
                'Orientación
                page.PageSetup.Orientation = Orientation.Landscape

                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE ASUNTOS EN TRÁMITE Y PROYECTOS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(4.3))
                currentTable.AddColumn(Unit.FromInch(1.5))
                currentTable.AddColumn(Unit.FromInch(2.5))
                currentTable.AddColumn(Unit.FromInch(2.5))

                'Heading row 
                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("NOMBRE Y DESCRIPCIÓN DEL ASUNTO Y/O PROYECTO")
                tRow(1).AddParagraph("FECHA DE PRESENTACIÓN")
                tRow(2).AddParagraph("JUSTIFICACIÓN Y/O SITUACIÓN")
                tRow(3).AddParagraph("OBSERVACIONES")

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

            Case "F.4"
                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE SELLOS AUTORIZADOS", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(6))
                currentTable.AddColumn(Unit.FromInch(1.4))

                'Heading row 
                tRow = currentTable.AddRow()

                tRow(0).AddParagraph("DESCRIPCIÓN")
                tRow(1).AddParagraph("SELLO")

                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data 
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                Next

                page.Add(currentTable)


            Case "I"
                'Headings 
                paragraph = page.AddParagraph("INFORME DE ACTIVIDADES", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(5))
                currentTable.AddColumn(Unit.FromInch(2.4))

                'Heading row 

                tRow = currentTable.AddRow()


                tRow(0).AddParagraph("NÚMERO DE ACTIVIDAD")
                tRow(1).AddParagraph("INFORME DE ACTIVIDADES")

                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data 
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                Next

                page.Add(currentTable)

            Case "II"
                'Headings 
                paragraph = page.AddParagraph("ORGANIGRAMA", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(7.4))

                'Heading row 

                tRow = currentTable.AddRow()


                tRow(0).AddParagraph("ORGANIGRAMA")

                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data 
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                Next

                page.Add(currentTable)

            Case "III"
                'Headings 
                paragraph = page.AddParagraph("FUNCIONES GENERALES", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(7.4))

                'Heading row 

                tRow = currentTable.AddRow()


                tRow(0).AddParagraph("FUNCION GENERAL")

                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data 
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                Next

                page.Add(currentTable)

            Case "IV"
                'Headings 
                paragraph = page.AddParagraph("RELACIÓN DE ANEXOS APLICABLES", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(1.7))
                currentTable.AddColumn(Unit.FromInch(2.4))
                currentTable.AddColumn(Unit.FromInch(1.7))
                currentTable.AddColumn(Unit.FromInch(1.7))

                'Heading row 

                tRow = currentTable.AddRow()


                tRow(0).AddParagraph("SECCIÓN")
                tRow(1).AddParagraph("CLAVE DEL ANEXO")
                tRow(2).AddParagraph("NOMBRE DEL ANEXO")
                tRow(3).AddParagraph("APLICA")

                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data 
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                    tRow.Cells(1).AddParagraph(dt(index)(3))
                    tRow.Cells(2).AddParagraph(dt(index)(4))
                    tRow.Cells(3).AddParagraph(dt(index)(6))
                Next

                page.Add(currentTable)

            Case "V"
                'Headings 
                paragraph = page.AddParagraph("PLAN MUNICIPAL DE DESARROLLO", "Heading1")
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
                currentTable.AddColumn(Unit.FromInch(7.4))

                'Heading row 

                tRow = currentTable.AddRow()


                tRow(0).AddParagraph("PLAN MUNICIPAL DE DESARROLLO")

                tRow.Format.Font.Bold = True
                tRow.Format.Alignment = ParagraphAlignment.Center

                'Data 
                For index = 0 To dt.Rows.Count - 1
                    tRow = currentTable.AddRow()
                    tRow.Cells(0).AddParagraph(dt(index)(2))
                Next

                page.Add(currentTable)

        End Select
        Return page

    End Function
End Class
