Imports MigraDoc.DocumentObjectModel
Imports MigraDoc.DocumentObjectModel.Tables

Partial Public Class WebForm1
    Public Function CreatePDF(r As List(Of Integer)) As Document

        Dim doc As New Document()
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




        Dim page = doc.AddSection()
        page.PageSetup.LeftMargin = Unit.FromInch(0.5)
        page.PageSetup.RightMargin = Unit.FromInch(0.5)
        page.PageSetup.TopMargin = Unit.FromInch(0.5)

        'Header
        Dim headerTable As New Table()
        headerTable.Borders.Width = 0.75
        headerTable.AddColumn(Unit.FromInch(4))
        headerTable.AddColumn(Unit.FromInch(1.4))
        headerTable.AddColumn(Unit.FromInch(2))
        headerTable.TopPadding = 4
        headerTable.BottomPadding = 4
        headerTable.LeftPadding = 4
        headerTable.RightPadding = 4

        Dim headerRow = headerTable.AddRow()
        headerRow.Shading.Color = Colors.LightGray
        headerRow.Cells(0).AddParagraph("GOBIERNO MUNICPAL DE SAN NICOLÁS DE LOS GARZA" & Environment.NewLine & " ADMINISTRACIÓN 2021 - 2024" & Environment.NewLine & "PROGRAMA DE ENTREGA-RECEPCIÓN PARA LA ADMINSITRACIÓN PÚBLICA MUNICIPAL" & Environment.NewLine & "RECURSOS FINANCIEROS" & Environment.NewLine & "ANEXO A.1")
        headerRow.Cells(0).Format.Alignment = ParagraphAlignment.Center
        headerRow.Cells(1).Shading.Color = Colors.White
        headerRow.Cells(1).Format.Alignment = ParagraphAlignment.Center
        Dim NLlogo = headerRow.Cells(1).AddParagraph.AddImage(AppDomain.CurrentDomain.BaseDirectory & "bin\Logo_NuevoLeon.png")
        NLlogo.Width = 64
        NLlogo.Height = 80
        'NLlogo.
        headerRow.Cells(2).AddParagraph("SECRETARÍA " & r(0) & Environment.NewLine & "DIRECCIÓN " & r(1))

        doc.LastSection.Add(headerTable)

        'Headings
        Dim paragraph = doc.LastSection.AddParagraph("PRESUPUESTO GLOBAL 2023", "Heading1")
        paragraph.Format.Borders.Width = 2.5
        paragraph.Format.Borders.Color = Colors.Black
        paragraph.Format.Borders.Distance = 3
        paragraph.Format.Shading.Color = Colors.Gray

        doc.LastSection.AddParagraph("SE ANEXA INFORMACIÓN" & Environment.NewLine & "PRESUPUESTO DE INGRESOS Y EGRESOS GLOBALES" & Environment.NewLine & "(MILES DE PESOS)" & Environment.NewLine, "Heading2")

        doc.LastSection.AddParagraph(" ")

        'Create first table
        Dim currentTable As New Table()
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
        Dim tRow = currentTable.AddRow()
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
        tRow(2).AddParagraph(r(2))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("PRESUPUESTO MODIFICADO")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(3))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("PRESUPUESTO DEVENGADO")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(4))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("PRESUPUESTO RECAUDADO")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(5))


        doc.LastSection.Add(currentTable)
        'Empty space
        doc.LastSection.AddParagraph(" " & Environment.NewLine & " ")

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
        tRow(2).AddParagraph(r(6))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("1ER. AMPLIACIÓN PRESUPUESTAL (SE ANEXA DOC.)")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(7))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("2DA. AMPLIACIÓN PRESUPUESTAL (SE ANEXA DOC.)")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(8))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("3ER. AMPLIACIÓN PRESUPUESTAL (SE ANEXA DOC.)")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(9))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("TOTAL AMPLIACIONES (2 + 3 + 4)")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(10))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("PRESUPUESTO MODIFICADO (SE ANEXA DOC.)")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(11))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("PRESUPUESTO COMPROMETIDO (SE ANEXA DOC.)")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(12))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("PRESUPUESTO DEVENGADO (SE ANEXA DOC.)")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(13))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("PRESUPUESTO EJERCIDO (SE ANEXA DOC.)")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(14))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("PRESUPUESTO EROGADO (SE ANEXA DOC.)")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(15))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("PRESUPUESTO CONSUMIDO (SE ANEXA DOC.) (7+8+9+10)")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(16))
        tRow = currentTable.AddRow()
        tRow(0).AddParagraph("PRESUPUESTO POR EJERCER OFICIAL (SE ANEXA DOC.) (6 -11)")
        tRow(1).AddParagraph("$")
        tRow(2).AddParagraph(r(17))

        doc.LastSection.Add(currentTable)

        Return doc
    End Function
End Class
