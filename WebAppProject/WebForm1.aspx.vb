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

Public Class WebForm1
    Inherits Page

    Private r As New List(Of Integer)({1000, 2000, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16})

    <Obsolete>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim pdfDoc = CreatePDF(r)
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
    Private Function ExportPDF()

    End Function

    Protected Sub generateReport_Click(sender As Object, e As EventArgs) Handles generateReport.Click

        Dim pdfDoc = CreatePDF(r)
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

    Protected Function CreatePDF(r As List(Of Integer)) As Document

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

        'Header
        Dim headerTable As New Table()
        headerTable.Borders.Width = 0.75
        headerTable.AddColumn(Unit.FromInch(3))
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
        currentTable.AddColumn(Unit.FromInch(4))
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
        currentTable.AddColumn(Unit.FromInch(4))
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