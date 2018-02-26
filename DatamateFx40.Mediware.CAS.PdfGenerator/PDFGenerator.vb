Imports Microsoft.VisualBasic
Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.tool.xml.css
Imports iTextSharp.tool.xml
Imports iTextSharp.tool.xml.html
Imports iTextSharp.tool.xml.pipeline.html
Imports iTextSharp.tool.xml.pipeline.css
Imports iTextSharp.tool.xml.pipeline.end
Imports iTextSharp.tool.xml.parser
Imports iTextSharp.text.pdf
Imports iTextSharp.text.html.simpleparser
Imports iTextSharp.tool.xml.pipeline
Imports System.Web
Imports System.Text.RegularExpressions
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Text
Imports iTextSharp.tool.xml.html.table

Public MustInherit Class PdfGeneratorProperties
    Protected Property HeaderData As String = ""
    Protected Property ContentData As String = ""
    Protected Property FooterData As String = ""
    Protected Property ReportHeader As String = ""
    ''' <summary>
    ''' It defines the page type of pdf document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Document Page Size</remarks>
    Public Property DocPageSize As DocumentPageSize = DocumentPageSize.A4
    ''' <summary>
    ''' Document page height in inch
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DocPageHeight As Double = 0
    ''' <summary>
    ''' document page width in inch
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DocPageWidth As Double = 0

    Protected Property PageHeight As Single = 0
    Protected Property PageWidth As Single = 0

    ''' <summary>
    ''' Document left margin in inch.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PageLeftMargin As Single = 0.75
    ''' <summary>
    ''' Document right margin in inch.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PageRightMargin As Single = 0.52
    ''' <summary>
    ''' Document Top Margin in Inch.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PageTopMargin As Single = 1.0
    ''' <summary>
    ''' Document Bottom margin in inch.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PageBottomMargin As Single = 1.44

    ''' <summary>
    ''' Css File Name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CssFileName As String = "" 'CommonStyles.css
    ''' <summary>
    ''' List of css
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Property CssFileNames As List(Of String)
    ''' <summary>
    ''' Css File Data
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Property CssFileData As String = ""

    ''' <summary>
    ''' Html File Name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SourceHtmlFileName As String = ""
    ''' <summary>
    ''' Html file in string format.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SourceHtmlString As String = ""
    ''' <summary>
    ''' List of Html files
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SourceHtmlStrings As List(Of String)

    ''' <summary>
    ''' true for landscape default false
    ''' </summary>
    ''' <value>True,False</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IsLandScape As Boolean = False

    Public Property IsPageNumberRequired As Boolean = True
    Public Property PageNumberPosition As PageNumberAlign = PageNumberAlign.Center
    Public Property IsStrokeRequired As Boolean = True

    Public Property IsWaterMarkRequired As Boolean = False
    Public Property WaterMarkContent As String = ""
    Public Property WaterMarkContentType As WarterMarkType = WarterMarkType.Text

    Public Property DebugMode As Boolean = False
    Public Property ErrorString As String = ""

    Public Property DirectPrint As Boolean = True

    Public Property PageRotation As RotationMode = RotationMode.Portrait

    Protected Property UnicodeFont As String = ""

    Public Property SeperateHeader As Boolean = False

    Public MustOverride Function GeneratePdfDocument() As Document
    Public MustOverride Function GeneratePdfDocument(ByVal bool As Boolean) As MemoryStream
    Public MustOverride Function GeneratePdfDocument(ByVal bool As Boolean, ByVal bool1 As Boolean) As Byte()

    Enum DocumentPageSize
        A4
        A5
        Custom
    End Enum

    Enum PageNumberAlign
        Left
        Center
        Right
    End Enum

    Enum WarterMarkType
        Image
        Text
    End Enum

    Enum RotationMode
        Portrait = 0
        Landscape = 90
        InvertedPortrait = 180
        Seascape = 270
    End Enum
End Class

Public Class PDFGenerator
    Inherits PdfGeneratorProperties

    Public Overrides Function GeneratePdfDocument() As Document
        ExtractHtmlData()
        Return CreatePdfDocument()
    End Function

    Public Overrides Function GeneratePdfDocument(ByVal bool As Boolean) As MemoryStream
        ExtractHtmlData()
        Dim _doc = CreatePdfDocument()

        Dim ms As New MemoryStream()
        Dim sw As New StreamWriter(ms)
        sw.Write(_doc)
        sw.Flush()
        ms.Position = 0

        Return ms
    End Function

    Public Overrides Function GeneratePdfDocument(ByVal bool As Boolean, ByVal bool1 As Boolean) As Byte()
        ExtractHtmlData()
        Dim _doc = CreatePdfDocument()

        Dim ms As New MemoryStream()
        Dim sw As New StreamWriter(ms)
        sw.Write(_doc)
        sw.Flush()
        ms.Position = 0

        Return ms.ToArray()
    End Function

    Private Sub ExtractHtmlData()
        Try
            Dim PrintFileStream As StreamReader
            Dim _FileData As String = ""
            If SourceHtmlFileName <> "" And SourceHtmlString = "" Then
                Dim FilePath = HttpContext.Current.Server.MapPath("~") & "/" & SourceHtmlFileName
                Dim fi = New FileInfo(FilePath)
                PrintFileStream = fi.OpenText()
                _FileData = HttpContext.Current.Server.HtmlDecode(PrintFileStream.ReadToEnd())
                PrintFileStream.Close()
            Else
                _FileData = SourceHtmlString
            End If

            SetProperties(_FileData)

            Dim _LstCssFiles As New List(Of String)

            _LstCssFiles = GetImageCssSrcFromString(_FileData)


            RemoveHiddenElements(_FileData)

            If _LstCssFiles.Count = 1 Then
                Me.CssFileName = _LstCssFiles(0)
            Else
                Me.CssFileNames = _LstCssFiles
            End If


            Dim sDelimStart As String = "<!--[HeaderStart]-->"
            Dim sDelimEnd As String = "<!--[HeaderEnd]-->"

            Dim nIndexStart As Integer = _FileData.IndexOf(sDelimStart) 'Find the first occurrence of f1
            Dim nIndexEnd As Integer = _FileData.IndexOf(sDelimEnd) 'Find the first occurrence of f2

            Dim HeaderHtmlData As String = ""
            If nIndexStart > -1 AndAlso nIndexEnd > -1 Then '-1 means the word was not found.
                HeaderHtmlData = Strings.Mid(_FileData, nIndexStart + sDelimStart.Length + 1, nIndexEnd - nIndexStart - sDelimStart.Length) 'Crop the text between
                'sSource = sSource.Remove(nIndexStart - sDelimStart.Length, nIndexEnd + sDelimEnd.Length)
            End If

            sDelimStart = "<!--[ContentStart]-->"
            sDelimEnd = "<!--[ContentEnd]-->"

            nIndexStart = _FileData.IndexOf(sDelimStart) 'Find the first occurrence of f1
            nIndexEnd = _FileData.IndexOf(sDelimEnd) 'Find the first occurrence of f2

            Dim ContentHtmlData As String = ""
            If nIndexStart > -1 AndAlso nIndexEnd > -1 Then '-1 means the word was not found.
                ContentHtmlData = Strings.Mid(_FileData, nIndexStart + sDelimStart.Length + 1, nIndexEnd - nIndexStart - sDelimStart.Length) 'Crop the text between
                'sSource = sSource.Remove(nIndexStart - sDelimStart.Length, nIndexEnd + sDelimEnd.Length)
            End If


            sDelimStart = "<!--[FooterStart]-->"
            sDelimEnd = "<!--[FooterEnd]-->"

            nIndexStart = _FileData.IndexOf(sDelimStart) 'Find the first occurrence of f1
            nIndexEnd = _FileData.IndexOf(sDelimEnd) 'Find the first occurrence of f2

            Dim FooterHtmlData As String = ""
            If nIndexStart > -1 AndAlso nIndexEnd > -1 Then '-1 means the word was not found.
                FooterHtmlData = Strings.Mid(_FileData, nIndexStart + sDelimStart.Length + 1, nIndexEnd - nIndexStart - sDelimStart.Length) 'Crop the text between
                'sSource = sSource.Remove(nIndexStart - sDelimStart.Length, nIndexEnd + sDelimEnd.Length)
            End If

            sDelimStart = "<!--[ReportHeaderStart]-->"
            sDelimEnd = "<!--[ReportHeaderEnd]-->"

            nIndexStart = _FileData.IndexOf(sDelimStart) 'Find the first occurrence of f1
            nIndexEnd = _FileData.IndexOf(sDelimEnd) 'Find the first occurrence of f2

            Dim ReportHeaderHtmlData As String = ""
            If nIndexStart > -1 AndAlso nIndexEnd > -1 Then '-1 means the word was not found.
                ReportHeaderHtmlData = Strings.Mid(_FileData, nIndexStart + sDelimStart.Length + 1, nIndexEnd - nIndexStart - sDelimStart.Length) 'Crop the text between
                'sSource = sSource.Remove(nIndexStart - sDelimStart.Length, nIndexEnd + sDelimEnd.Length)
            End If

            sDelimStart = "<style type=""text/css"">"
            sDelimEnd = "</style>"

            nIndexStart = _FileData.IndexOf(sDelimStart) 'Find the first occurrence of f1
            nIndexEnd = _FileData.IndexOf(sDelimEnd) 'Find the first occurrence of f2

            Dim CssHtmlData As String = ""
            If nIndexStart > -1 AndAlso nIndexEnd > -1 Then '-1 means the word was not found.
                CssHtmlData = Strings.Mid(_FileData, nIndexStart + sDelimStart.Length + 1, nIndexEnd - nIndexStart - sDelimStart.Length) 'Crop the text between
                'sSource = sSource.Remove(nIndexStart - sDelimStart.Length, nIndexEnd + sDelimEnd.Length)
            End If


            Me.ContentData = ContentHtmlData
            Me.HeaderData = HeaderHtmlData
            Me.FooterData = FooterHtmlData
            Me.ReportHeader = ReportHeaderHtmlData

            Me.CssFileData = CssHtmlData

            If ReportHeader.Trim.Length > 0 Then
                Me.SeperateHeader = True
            Else
                Me.SeperateHeader = False
            End If

        Catch ex As Exception
            Me.ErrorString = ex.Message
        End Try

    End Sub

    Public Sub GenerateAndDownloadPdf(ByVal FileName As String, ByVal IsDownload As Boolean)
        Try
            HttpContext.Current.Response.ContentType = "application/pdf"
            If IsDownload Then
                HttpContext.Current.Response.AddHeader("content-disposition", "attachment;filename=" & FileName)
            Else
                HttpContext.Current.Response.AddHeader("content-disposition", "inline;filename=" & FileName)
            End If

            HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache)

            HttpContext.Current.Response.ContentEncoding = Encoding.UTF8

            HttpContext.Current.Response.Write(GeneratePdfDocument(True))

            HttpContext.Current.Response.HeaderEncoding = Encoding.UTF8

            HttpContext.Current.ApplicationInstance.CompleteRequest()
        Catch ex As Exception
            Me.ErrorString = ex.Message
            Me.DebugMode = DebugMode
        End Try
    End Sub

    Private Function GetImageCssSrcFromString(ByRef HtmlString As String) As List(Of String)
        Try

            Dim webpath As String = HttpContext.Current.Request.Url.AbsoluteUri.Replace(HttpContext.Current.Request.Url.PathAndQuery, "") & HttpContext.Current.Request.ApplicationPath

            webpath = webpath.Trim & "/"

            If Me.WaterMarkContentType = WarterMarkType.Image Then
                Me.WaterMarkContent = webpath & Me.WaterMarkContent.TrimStart("/")
            End If

            Dim imgSrcMatches = Regex.Matches(HtmlString, String.Format("<\s*img\s*src\s*=\s*{0}\s*([^{0}]+)\s*{0}", """"), RegexOptions.CultureInvariant Or RegexOptions.IgnoreCase Or RegexOptions.Multiline)

            Dim imgSrcs As New List(Of String)()
            For Each match As Match In imgSrcMatches
                imgSrcs.Add(match.Groups(1).Value)
            Next

            imgSrcMatches = Regex.Matches(HtmlString, String.Format("<\s*img\s*src\s*=\s*{0}\s*([^{0}]+)\s*{0}", "'"), RegexOptions.CultureInvariant Or RegexOptions.IgnoreCase Or RegexOptions.Multiline)

            For Each match As Match In imgSrcMatches
                imgSrcs.Add(match.Groups(1).Value)
            Next

            For Each img In imgSrcs.Distinct().ToList()
                If Not img.StartsWith("http:") And Not img.StartsWith("data:") And Not img.StartsWith("https:") Then
                    Dim newImgPath = ""
                    If img.StartsWith("../") Then
                        Dim newimg = String.Copy(img)
                        newimg = newimg.Remove(0, 3).Trim()
                        newImgPath = webpath & newimg
                    Else
                        newImgPath = webpath & img
                    End If

                    HtmlString = HtmlString.Replace(img, newImgPath)
                End If
            Next

            Dim linkHrefMatches = Regex.Matches(HtmlString, String.Format("<\s*link\s*href\s*=\s*{0}\s*([^{0}]+)\s*{0}", """"), RegexOptions.CultureInvariant Or RegexOptions.IgnoreCase Or RegexOptions.Multiline)


            Dim linkHrefs As New List(Of String)()
            For Each match As Match In linkHrefMatches
                linkHrefs.Add(match.Groups(1).Value)
            Next


            linkHrefMatches = Regex.Matches(HtmlString, String.Format("<\s*link\s*href\s*=\s*{0}\s*([^{0}]+)\s*{0}", "'"), RegexOptions.CultureInvariant Or RegexOptions.IgnoreCase Or RegexOptions.Multiline)

            For Each match As Match In linkHrefMatches
                linkHrefs.Add(match.Groups(1).Value)
            Next

            For Each linkhref In linkHrefs.Distinct().ToList()
                Dim newLinkPath = ""
                newLinkPath = HttpContext.Current.Server.MapPath("~") & "/" & linkhref.TrimStart("../")
                linkHrefs.Remove(linkhref)
                linkHrefs.Add(newLinkPath)
            Next

            Return linkHrefs
        Catch ex As Exception
            Me.ErrorString = ex.Message
            Return Nothing
        End Try
    End Function

    Private Function CreatePdfDocument() As Document
        Try
            Dim _PageSize As New Object()

            Select Case DocPageSize
                Case DocumentPageSize.A4
                    _PageSize = PageSize.A4
                Case DocumentPageSize.A5
                    _PageSize = PageSize.A5
                Case DocumentPageSize.Custom
                    Me.PageHeight = DocPageHeight * 72
                    Me.PageWidth = DocPageWidth * 72

                    _PageSize = New Rectangle(PageWidth, PageHeight)
                Case Else
                    _PageSize = PageSize.A4
            End Select

            PageLeftMargin = PageLeftMargin * 72
            PageRightMargin = PageRightMargin * 72
            PageTopMargin = PageTopMargin * 72
            PageBottomMargin = PageBottomMargin * 72

            Dim _document As New Document(_PageSize, PageLeftMargin, PageRightMargin, PageTopMargin, PageBottomMargin)


            If IsLandScape Then
                _document.SetPageSize(_PageSize.Rotate())
                'Select Case DocPageSize
                '    Case DocumentPageSize.A4
                '        _document.SetPageSize(PageSize.A4.Rotate()) '-- landscape
                '    Case DocumentPageSize.A5
                '        _document.SetPageSize(PageSize.A5.Rotate()) '-- landscape
                '    Case DocumentPageSize.Custom
                '        _document.SetPageSize(_PageSize.Rotate())
                '    Case Else
                '        _document.SetPageSize(PageSize.A4.Rotate()) '-- landscape
                'End Select
            End If


           
         
            '_document.SetAccessibleAttribute
            ' _document.
            Dim _Pdfwriter As PdfWriter = PdfWriter.GetInstance(_document, HttpContext.Current.Response.OutputStream)

            _document.AddAuthor("© Datamate Info Solutions ®")
            _document.AddCreationDate()
            _document.AddProducer()
            _document.AddSubject("Dmate pdf Print")

            Dim _pdfevent = New PdfPageClass(Me)

            _pdfevent.HeaderData = HeaderData.Trim
            _pdfevent.FooterData = FooterData.Trim
            _pdfevent.ReportHeader = ReportHeader.Trim
            _pdfevent.IsLandScape = IsLandScape

            Select Case PageRotation
                Case RotationMode.Portrait
                    _pdfevent.SetOrientation(PdfPage.PORTRAIT)
                Case RotationMode.Landscape
                    _pdfevent.SetOrientation(PdfPage.LANDSCAPE)
                Case RotationMode.InvertedPortrait
                    _pdfevent.SetOrientation(PdfPage.INVERTEDPORTRAIT)
                Case RotationMode.Seascape
                    _pdfevent.SetOrientation(PdfPage.SEASCAPE)
                Case Else
                    _pdfevent.SetOrientation(PdfPage.PORTRAIT)
            End Select



            GetCssDataFromFile(CssFileNames, Me.CssFileData, CssFileName)

            _pdfevent.CssFileData = Me.CssFileData

            Dim msCssStream As New MemoryStream()
            Dim _swCss As New StreamWriter(msCssStream)
            _swCss.Write(CssFileData)
            _swCss.Flush()
            msCssStream.Position = 0

            ''''Adding ReportHeader
            Dim header As ElementList

            header = _pdfevent.ParseHtmltoElements(HeaderData.Trim)

            Dim rptheaderTbl As New PdfPTable(1)
            rptheaderTbl.SetTotalWidth({_document.PageSize.Width - PageLeftMargin - PageRightMargin})

            Dim RptCell As New PdfPCell()
            RptCell.Border = 0

            For Each e As IElement In header
                RptCell.AddElement(e)
            Next

            rptheaderTbl.AddCell(RptCell)

            _pdfevent.ReportTable = rptheaderTbl

            ''''Adding Header
            Dim HeaderDataHtml = ReportHeader.Trim & HeaderData.Trim

            header = _pdfevent.ParseHtmltoElements(HeaderDataHtml)

            Dim headerTbl As New PdfPTable(1)
            headerTbl.SetTotalWidth({_document.PageSize.Width - PageLeftMargin - PageRightMargin})

            Dim cell As New PdfPCell()
            cell.Border = 0

            For Each e As IElement In header
                cell.AddElement(e)
            Next

            headerTbl.AddCell(cell)

            _pdfevent.HeaderTable = headerTbl

            ''''Adding Footer

            header = _pdfevent.ParseHtmltoElements(FooterData.Trim)

            Dim footerTbl As New PdfPTable(1)
            footerTbl.SetTotalWidth({_document.PageSize.Width - PageLeftMargin - PageRightMargin})

            Dim FooterCell As New PdfPCell()
            FooterCell.Border = 0

            For Each e As IElement In header
                FooterCell.AddElement(e)
            Next

            footerTbl.AddCell(FooterCell)

            _pdfevent.FooterTable = footerTbl

            _document.SetMargins(PageLeftMargin, PageRightMargin, PageTopMargin + headerTbl.TotalHeight(), PageBottomMargin + footerTbl.TotalHeight())

            _Pdfwriter.PageEvent = _pdfevent

            _pdfevent.NewPage(_document)

            ''''' adding Content.

            Dim _LitContent As New Literal()

            _LitContent.Text = ContentData


            Dim sw As New StringWriter()
            Dim hw As New HtmlTextWriter(sw)
            _LitContent.RenderControl(hw)
            Dim sr As New StringReader(sw.ToString())

            Dim cssResolver = New StyleAttrCSSResolver()

            Dim cssFile As New Object()

            cssFile = XMLWorkerHelper.GetCSS(msCssStream)
            cssResolver.AddCss(cssFile)
            msCssStream.Dispose()
            msCssStream.Close()



            _document.Open()


            Dim _UnicodeFontFactory = New UnicodeFontFactory(UnicodeFont)

            Dim cssAppliers As CssAppliers = New CssAppliersImpl(_UnicodeFontFactory)

            Dim HtmlContext As New HtmlPipelineContext(cssAppliers)

            HtmlContext.SetTagFactory(Tags.GetHtmlTagProcessorFactory)

            Dim tagProcessors = DirectCast(Tags.GetHtmlTagProcessorFactory(), DefaultTagProcessorFactory)
            tagProcessors.RemoveProcessor(HTML.Tag.IMG)
            ' remove the default processor
            tagProcessors.AddProcessor(HTML.Tag.IMG, New CustomImageTagProcessor())
            tagProcessors.AddProcessor(HTML.Tag.TD, New TableDataProcessor())
            ' use our new processor

            HtmlContext.SetAcceptUnknown(True).AutoBookmark(True).SetTagFactory(tagProcessors)

            Dim Pipeline As IPipeline = New CssResolverPipeline(cssResolver, New HtmlPipeline(HtmlContext, New PdfWriterPipeline(_document, _Pdfwriter)))

            Dim Worker As XMLWorker = New XMLWorker(Pipeline, True)

            Dim Parser As XMLParser = New XMLParser(True, Worker, Encoding.UTF8)

            Dim htmlinput = GenerateStreamFromString(sw.ToString())

            Parser.Parse(htmlinput, Encoding.UTF8)

            Parser.Flush()




            If DirectPrint Then
                'writer.AddJavaScript("this.print(true);", False)
                Dim action As New PdfAction(PdfAction.PRINTDIALOG)
                _Pdfwriter.SetOpenAction(action)
            End If

            

            _document.Close()
            _UnicodeFontFactory.DeleteFont()


            Return _document

        Catch ex As Exception
            Me.ErrorString = ex.Message
            Return Nothing
        End Try
    End Function

    Private Sub GetCssDataFromFile(ByVal CssFileNames As List(Of String), ByRef ReturnString As String, Optional ByVal CssFileDataName As String = "")
        Try
            Dim PrintFileStream As StreamReader = Nothing
            Dim _sbCss As New StringBuilder()
            If CssFileDataName = "" Then
                For Each cssfile In CssFileNames
                    Dim fi = New FileInfo(cssfile)
                    PrintFileStream = fi.OpenText()
                    _sbCss.AppendLine(HttpContext.Current.Server.HtmlDecode(PrintFileStream.ReadToEnd()))
                    PrintFileStream.Dispose()
                    PrintFileStream.Close()
                Next
            Else
                Dim fi = New FileInfo(CssFileDataName)
                PrintFileStream = fi.OpenText()
                _sbCss.AppendLine(HttpContext.Current.Server.HtmlDecode(PrintFileStream.ReadToEnd()))

                PrintFileStream.Dispose()
                PrintFileStream.Close()
            End If




            ReturnString = ReturnString & _sbCss.ToString()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetProperties(ByVal HtmlString As String)
        Try
            Dim pageSetingMatches = Regex.Matches(HtmlString, String.Format("<\s*PageSettings\s*data\s*=\s*{0}\s*([^{0}]+)\s*{0}", """"), RegexOptions.CultureInvariant Or RegexOptions.IgnoreCase Or RegexOptions.Multiline Or RegexOptions.IgnorePatternWhitespace)
            Dim pageKeys As New List(Of String)()
            For Each match As Match In pageSetingMatches
                pageKeys.Add(match.Groups(1).Value.ToLower())
            Next

            If pageKeys.Count > 0 Then
                Dim sKeyString As String = ""

                If pageKeys.Count = 1 Then
                    sKeyString = pageKeys(0)
                End If

                Dim nChar() As Char = {";", ":"}
                Dim sSplitval = sKeyString.Split(nChar)
                Dim dData As New Dictionary(Of String, Object)
                '''''''''''''''''' sSplitval contains all keys and values 

                For i As Integer = 0 To sSplitval.Length - 2 Step 2
                    dData.Add(sSplitval(i), sSplitval(i + 1))
                Next

                If dData.Keys.Contains("pagesize") Then
                    Select Case dData("pagesize")
                        Case "a4"
                            Me.DocPageSize = DocumentPageSize.A4
                        Case "a5"
                            Me.DocPageSize = DocumentPageSize.A5
                        Case "custom"
                            Me.DocPageSize = DocumentPageSize.Custom
                    End Select
                End If

                If Me.DocPageSize = DocumentPageSize.Custom Then
                    If dData.Keys.Contains("pagewidth") Then
                        Me.DocPageWidth = Val(dData("pagewidth").ToString().Trim())
                    Else
                        Throw New Exception("Set Page Width for Custom Page.")
                    End If

                    If dData.Keys.Contains("pageheight") Then
                        Me.DocPageHeight = Val(dData("pageheight").ToString().Trim())
                    Else
                        Throw New Exception("Set Page Height for Custom Page.")
                    End If
                End If

                If dData.Keys.Contains("margin") Then
                    Dim sMargins = dData("margin").ToString().Trim().Split(" ")
                    If sMargins.Length = 1 Then
                        Me.PageLeftMargin = Convert.ToSingle(sMargins(0).Trim)
                        Me.PageRightMargin = Convert.ToSingle(sMargins(0).Trim)
                        Me.PageTopMargin = Convert.ToSingle(sMargins(0).Trim)
                        Me.PageBottomMargin = Convert.ToSingle(sMargins(0).Trim)
                    Else
                        Me.PageLeftMargin = Convert.ToSingle(sMargins(0).Trim)
                        Me.PageRightMargin = Convert.ToSingle(sMargins(1).Trim)
                        Me.PageTopMargin = Convert.ToSingle(sMargins(2).Trim)
                        Me.PageBottomMargin = Convert.ToSingle(sMargins(3).Trim)
                    End If

                End If

                If dData.Keys.Contains("landscape") Then
                    If dData("landscape") = "true" Then
                        Me.IsLandScape = True
                    Else
                        Me.IsLandScape = False
                    End If
                End If

                If dData.Keys.Contains("watermark") Then
                    If dData("watermark") = "true" Then
                        Me.IsWaterMarkRequired = True
                    Else
                        Me.IsWaterMarkRequired = False
                    End If
                End If

                If Me.IsWaterMarkRequired Then
                    If dData.Keys.Contains("watermarktype") Then
                        Select Case dData("watermarktype")
                            Case "image"
                                Me.WaterMarkContentType = WarterMarkType.Image
                            Case "text"
                                Me.WaterMarkContentType = WarterMarkType.Text
                        End Select
                    End If

                    If dData.Keys.Contains("watermarkcontent") Then
                        Me.WaterMarkContent = dData("watermarkcontent")
                    End If
                End If

                If dData.Keys.Contains("pagenumber") Then
                    If dData("pagenumber") = "true" Then
                        Me.IsPageNumberRequired = True
                    Else
                        Me.IsPageNumberRequired = False
                    End If
                End If

                If Me.IsPageNumberRequired Then
                    If dData.Keys.Contains("pagenumberposition") Then
                        Select Case dData("pagenumberposition")
                            Case "center"
                                Me.PageNumberPosition = PageNumberAlign.Center
                            Case "left"
                                Me.PageNumberPosition = PageNumberAlign.Left
                            Case "right"
                                Me.PageNumberPosition = PageNumberAlign.Right
                            Case Else
                                Me.PageNumberPosition = PageNumberAlign.Center
                        End Select
                    End If
                End If

                If dData.Keys.Contains("linerequired") Then
                    If dData("linerequired") = "true" Then
                        Me.IsStrokeRequired = True
                    Else
                        Me.IsStrokeRequired = False
                    End If
                End If

                If dData.Keys.Contains("debug") Then
                    If dData("debug") = "true" Then
                        Me.DebugMode = True
                    Else
                        Me.DebugMode = False
                    End If
                End If


                If dData.Keys.Contains("directprint") Then
                    If dData("directprint") = "true" Then
                        Me.DirectPrint = True
                    Else
                        Me.DirectPrint = False
                    End If
                End If

                If dData.Keys.Contains("rotation") Then
                    Me.PageRotation = Val(dData("rotation"))
                End If

                If dData.Keys.Contains("unicodefont") Then
                    Me.UnicodeFont = HttpContext.Current.Server.MapPath("~") & "/" & dData("unicodefont").TrimStart("../")
                End If

              

            End If

        Catch ex As Exception
            Me.ErrorString = ex.Message
        End Try
    End Sub

    Private Sub RemoveHiddenElements(ByRef _FileData As String)
        Try
            ''''-- Remove Span -----'''''
            Dim sPatten0 = String.Format("<span[^>]*data-value='?""?{0}""'?[^>]*>.*?.</span>", "0")
            Dim sPatten01 = String.Format("<span[^>]*data-value='?""?{0}""'?[^>]*>.*?.</span>", "0.0")
            Dim sPatten02 = String.Format("<span[^>]*data-value='?""?{0}""'?[^>]*>.*?.</span>", "0.00")
            Dim sPatten03 = String.Format("<span[^>]*data-value='?""?{0}""'?[^>]*>.*?.</span>", "0.000")
            Dim sPatten04 = String.Format("<span[^>]*data-value='?""?{0}""'?[^>]*>.*?.</span>", "0.0000")

            Dim regex = New Regex(sPatten0, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten01, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten02, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten03, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten04, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            '''''''''''''''''''''''''''''

            ''''-- Remove div -----'''''
            sPatten0 = String.Format("<div[^>]*data-value='?""?{0}""'?[^>]*>.*?.</div>", "0")
            sPatten01 = String.Format("<div[^>]*data-value='?""?{0}""'?[^>]*>.*?.</div>", "0.0")
            sPatten02 = String.Format("<div[^>]*data-value='?""?{0}""'?[^>]*>.*?.</div>", "0.00")
            sPatten03 = String.Format("<div[^>]*data-value='?""?{0}""'?[^>]*>.*?.</div>", "0.000")
            sPatten04 = String.Format("<div[^>]*data-value='?""?{0}""'?[^>]*>.*?.</div>", "0.0000")

            regex = New Regex(sPatten0, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten01, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten02, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten03, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten04, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            '''''''''''''''''''''''''''''

            ''''-- Remove tr -----'''''
            sPatten0 = String.Format("<tr[^>]*data-value='?""?{0}""'?[^>]*>.*?.</tr>", "0")
            sPatten01 = String.Format("<tr[^>]*data-value='?""?{0}""'?[^>]*>.*?.</tr>", "0.0")
            sPatten02 = String.Format("<tr[^>]*data-value='?""?{0}""'?[^>]*>.*?.</tr>", "0.00")
            sPatten03 = String.Format("<tr[^>]*data-value='?""?{0}""'?[^>]*>.*?.</tr>", "0.000")
            sPatten04 = String.Format("<tr[^>]*data-value='?""?{0}""'?[^>]*>.*?.</tr>", "0.0000")

            regex = New Regex(sPatten0, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten01, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten02, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten03, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten04, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            '''''''''''''''''''''''''''''


            ''''-- Remove table -----'''''
            sPatten0 = String.Format("<table[^>]*data-value='?""?{0}""'?[^>]*>.*?.</table>", "0")
            sPatten01 = String.Format("<table[^>]*data-value='?""?{0}""'?[^>]*>.*?.</table>", "0.0")
            sPatten02 = String.Format("<table[^>]*data-value='?""?{0}""'?[^>]*>.*?.</table>", "0.00")
            sPatten03 = String.Format("<table[^>]*data-value='?""?{0}""'?[^>]*>.*?.</table>", "0.000")
            sPatten04 = String.Format("<table[^>]*data-value='?""?{0}""'?[^>]*>.*?.</table>", "0.0000")

            regex = New Regex(sPatten0, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten01, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten02, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten03, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            regex = New Regex(sPatten04, RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline Or RegexOptions.IgnorePatternWhitespace)
            _FileData = regex.Replace(_FileData, "")

            '''''''''''''''''''''''''''''

        Catch ex As Exception
            Me.ErrorString = ex.Message
        End Try
    End Sub


    Public Sub GenerateDefaultPDF()
        Try
            Using document = New Document(PageSize.A4, 30, 30, 30, 30)
                Dim writer = PdfWriter.GetInstance(document, HttpContext.Current.Response.OutputStream)
                Dim worker = XMLWorkerHelper.GetInstance()

                Dim sStr = "<p>&diams; &larr;  &darr; &harr; &uarr; &rarr;&oslash; &euro; &copy;</p>"

                Dim htmlInput As Stream = GenerateStreamFromString(sStr)

                document.Open()
                worker.ParseXHtml(writer, document, htmlInput, Nothing, Encoding.UTF8, New UnicodeFontFactory())

                document.Close()
                Dim FileName = "Defult.pdf"
                HttpContext.Current.Response.ContentType = "application/pdf"

                HttpContext.Current.Response.AddHeader("content-disposition", "inline;filename=" & FileName)

                HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache)

                HttpContext.Current.Response.ContentEncoding = Encoding.UTF8

                HttpContext.Current.Response.Write(document)

                HttpContext.Current.Response.HeaderEncoding = Encoding.UTF8

                HttpContext.Current.ApplicationInstance.CompleteRequest()
            End Using


            'HttpContext.Current.Response.Clear()
            'HttpContext.Current.Response.Charset = ""
            'HttpContext.Current.Response.ContentType = "application/pdf"
            'Dim strFileName As String = "123" + ".pdf"
            'HttpContext.Current.Response.AddHeader("Content-Disposition", Convert.ToString("inline; filename=") & strFileName)

            'Dim outXml As String = "<html><body style='font-family: MYSYLFAEN;'>" & vbCr & vbLf & "            <meta http-equiv='Content-Type' content='text/html; charset=UTF-8'/>" & vbCr & vbLf & "            eng-abc_arm-աբգ_rus-абв_eng-xyz <p>&diams; &larr;  &darr; &harr; &uarr; &rarr;&oslash; &euro; &copy;</p></body></html>"

            'Dim memStream As New MemoryStream()

            'Dim xmlString As TextReader = New StringReader(outXml)

            'Using document As New Document()
            '    Dim writer As PdfWriter = PdfWriter.GetInstance(document, memStream)
            '    document.SetPageSize(iTextSharp.text.PageSize.A4)
            '    document.Open()

            '    FontFactory.Register("C:/Windows/Fonts/sylfaen.ttf", "MYSYLFAEN")

            '    Dim byteArray As Byte() = System.Text.Encoding.UTF8.GetBytes(outXml)
            '    Dim ms As New MemoryStream(byteArray)

            '    XMLWorkerHelper.GetInstance().ParseXHtml(writer, document, ms, System.Text.Encoding.UTF8)

            '    document.Close()
            'End Using

            'HttpContext.Current.Response.BinaryWrite(memStream.ToArray())
            'HttpContext.Current.Response.[End]()
            'HttpContext.Current.Response.Flush()
        Catch ex As Exception

        End Try
    End Sub

    Public Shared Function GenerateStreamFromString(ByVal s As String) As Stream
        Dim stream As New MemoryStream()
        Dim writer As New StreamWriter(stream)
        writer.Write(s)
        writer.Flush()
        stream.Position = 0
        Return stream
    End Function
End Class

Friend Class UnicodeFontFactory
    Inherits FontFactoryImp
    Private FontPath As String = ""

    Private _baseFont As BaseFont

    Public Sub New()
        _baseFont = BaseFont.CreateFont(FontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
    End Sub

    Public Sub New(ByVal sFonthPath As String)

        If File.Exists(sFonthPath) Then
            Me.FontPath = sFonthPath
            _baseFont = BaseFont.CreateFont(FontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
        Else
            Me.FontPath = ""
        End If

    End Sub

    Public Overrides Function GetFont(ByVal fontname As String, ByVal encoding As String, ByVal embedded As Boolean, ByVal size As Single, ByVal style As Integer, ByVal color As BaseColor, _
     ByVal cached As Boolean) As Font
        Return New Font(_baseFont, size, style, color)
    End Function

    Public Sub DeleteFont()
        Me._baseFont = Nothing
        Me.FontPath = ""
    End Sub


End Class

Friend Class PdfPageClass
    Inherits PdfPageEventHelper

    Dim bf As pdf.BaseFont
    Dim fnt As Font
    Dim template As PdfTemplate
    Dim cb As PdfContentByte

    Dim pagecount As Integer = 0
    Dim npagecount As Integer = 0
    Dim TotPagno As Integer = 0
    Dim bHeaderDrawn As Boolean = False

    'Protected LandscapeMode As [Boolean] = False

    Public Property HeaderData As String = ""
    Public Property ContentData As String = ""
    Public Property FooterData As String = ""
    Public Property ReportHeader As String = ""

    Public Property HeaderTable As PdfPTable
    Public Property FooterTable As PdfPTable
    Public Property ReportTable As PdfPTable


    Public Property CssName As String = "" 'CommonStyles.css
    Public Property CssFileData As String = ""
    Public Property IsLandScape As Boolean = False

    Protected header As ElementList

    Public Property PageLeftMargin As Single = 30
    Public Property PageRightMargin As Single = 30
    Public Property PageTopMargin As Single = 10
    Public Property PageBottomMargin As Single = 70

    Public Property IsPageNumberRequired As Boolean = False
    Public Property PageNumberPosition As PdfGeneratorProperties.PageNumberAlign = PdfGeneratorProperties.PageNumberAlign.Center
    Public Property IsStrokeRequired As Boolean = True

    Public Property IsWaterMarkRequired As Boolean = False
    Public Property WaterMarkContent As String = ""
    Protected Property WaterMarkContentType As PdfGeneratorProperties.WarterMarkType = PdfGeneratorProperties.WarterMarkType.Text

    Protected orientation As PdfNumber = PdfPage.PORTRAIT
    Protected Property SeprateHeader As Boolean = False
    Protected myNewPage As [Boolean] = False





    Sub New(ByRef pdfgen As PDFGenerator)

        Me.PageLeftMargin = pdfgen.PageLeftMargin
        Me.PageRightMargin = pdfgen.PageRightMargin
        Me.PageTopMargin = pdfgen.PageTopMargin
        Me.PageBottomMargin = pdfgen.PageBottomMargin

        Me.IsPageNumberRequired = pdfgen.IsPageNumberRequired
        Me.PageNumberPosition = pdfgen.PageNumberPosition
        Me.IsStrokeRequired = pdfgen.IsStrokeRequired

        Me.IsWaterMarkRequired = pdfgen.IsWaterMarkRequired
        Me.WaterMarkContent = pdfgen.WaterMarkContent
        Me.WaterMarkContentType = pdfgen.WaterMarkContentType
        Me.SeprateHeader = pdfgen.SeperateHeader

    End Sub


    Public Function ParseHtmltoElements(ByVal HtmlData As String) As ElementList

        Dim _FileDatacss As String = ""
        _FileDatacss = CssFileData
        
        Return XMLWorkerHelper.ParseToElementList(HtmlData, _FileDatacss)
    End Function

    Public Overrides Sub OnOpenDocument(ByVal writer As PdfWriter, ByVal document As Document)
        MyBase.OnOpenDocument(writer, document)
        Try

            bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)
            fnt = New Font(bf, 7.5)
            cb = writer.DirectContent
            template = cb.CreateTemplate(50, 50)
        Catch de As DocumentException
        Catch ioe As System.IO.IOException
        End Try
    End Sub

    Public Overrides Sub OnStartPage(ByVal writer As PdfWriter, ByVal document As Document)
        'Dim genFun As New GenFunction()

        MyBase.OnStartPage(writer, document)

        writer.AddPageDictEntry(PdfName.ROTATE, orientation)
        Try
            TotPagno = TotPagno + 1

            If Not myNewPage And SeprateHeader Then
                bHeaderDrawn = True
                HeaderTable = ReportTable
                document.SetMargins(PageLeftMargin, PageRightMargin, PageTopMargin + ReportTable.TotalHeight(), PageBottomMargin + FooterTable.TotalHeight())
            End If

            HeaderTable.WriteSelectedRows(0, 20, document.Left, document.PageSize.Top - PageTopMargin, writer.DirectContent)
            FooterTable.WriteSelectedRows(0, 20, document.Left, document.PageSize.Bottom + PageBottomMargin + FooterTable.TotalHeight(), writer.DirectContent)

            If IsWaterMarkRequired Then
                Select Case WaterMarkContentType
                    Case PdfGeneratorProperties.WarterMarkType.Image

                        Dim img As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(WaterMarkContent) '
                        img.SetAbsolutePosition((document.PageSize.Width - img.ScaledWidth) / 2, (document.PageSize.Height - img.ScaledHeight) / 2)
                        img.GrayFill = 1
                        img.RotationDegrees = 0

                        Dim content As PdfContentByte = writer.DirectContentUnder
                        content.AddImage(img)

                    Case PdfGeneratorProperties.WarterMarkType.Text
                        Dim fontSize As Single = 80
                        Dim xPosition As Single = 300
                        Dim yPosition As Single = 400
                        Dim angle As Single = 45
                        Try
                            Dim under As PdfContentByte = writer.DirectContentUnder
                            Dim baseFont__1 As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.EMBEDDED)
                            under.BeginText()
                            under.SetColorFill(BaseColor.LIGHT_GRAY)
                            under.SetFontAndSize(baseFont__1, fontSize)
                            under.ShowTextAligned(PdfContentByte.ALIGN_CENTER, WaterMarkContent, xPosition, yPosition, angle)
                            under.EndText()
                        Catch ex As Exception
                            Throw New Exception(ex.Message)
                        End Try
                End Select
            End If

        Catch ex As Exception
            ' genFun.AddToLogFile("Method Name : " & System.Reflection.MethodBase.GetCurrentMethod().Name & " , Exception Name: " & ex.Message & " Exception Path : " & HttpContext.Current.Request.RawUrl & " Exception Source : " & ex.Source)
        End Try
    End Sub

    Public Overrides Sub OnEndPage(ByVal writer As PdfWriter, ByVal document As Document)
        MyBase.OnEndPage(writer, document)

        pagecount = pagecount + 1

        npagecount = pagecount

        Dim pageN As Integer = pagecount
        Dim text As String = "Page  " & pageN & "  of  "
        Dim len As Single = bf.GetWidthPoint(text, 6)
        Dim pageSize As Rectangle = document.PageSize

        If IsPageNumberRequired Then

            cb.SetRGBColorFill(0, 0, 0)

            cb.BeginText()
            cb.SetFontAndSize(bf, 7.5)
            cb.SetTextMatrix(pageSize.GetRight((document.PageSize.Width) / 2 + text.Length), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight)))

            If Not IsLandScape Then
                Select Case PageNumberPosition
                    Case PdfGeneratorProperties.PageNumberAlign.Center
                        cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, pageSize.GetRight((document.PageSize.Width) / 2 + text.Length), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)), 0) '320,38
                    Case PdfGeneratorProperties.PageNumberAlign.Left
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, pageSize.GetLeft(pageSize.GetLeft(PageLeftMargin)), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)), 0) '55,38
                    Case PdfGeneratorProperties.PageNumberAlign.Right
                        cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, text, pageSize.GetRight(PageRightMargin + 5), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)), 0) '55,38
                End Select

            Else

                Select Case PageNumberPosition
                    Case PdfGeneratorProperties.PageNumberAlign.Center
                        cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, pageSize.GetRight((document.PageSize.Width) / 2 + text.Length), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)), 0)
                    Case PdfGeneratorProperties.PageNumberAlign.Left
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, pageSize.GetLeft(PageLeftMargin), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)), 0)
                    Case PdfGeneratorProperties.PageNumberAlign.Right
                        cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, text, pageSize.GetRight(PageRightMargin + 5), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)), 0)
                End Select
            End If
            cb.EndText()

            If Not IsLandScape Then
                Select Case PageNumberPosition
                    Case PdfGeneratorProperties.PageNumberAlign.Center
                        cb.AddTemplate(template, pageSize.GetRight(((document.PageSize.Width) / 2) - text.Length), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)))
                    Case PdfGeneratorProperties.PageNumberAlign.Left
                        cb.AddTemplate(template, pageSize.GetLeft(PageLeftMargin + text.Length + 32), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)))
                    Case PdfGeneratorProperties.PageNumberAlign.Right
                        cb.AddTemplate(template, pageSize.GetRight(PageRightMargin), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)))
                End Select
            Else
                Select Case PageNumberPosition
                    Case PdfGeneratorProperties.PageNumberAlign.Center
                        cb.AddTemplate(template, pageSize.GetRight(((document.PageSize.Width) / 2) - text.Length), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)))
                    Case PdfGeneratorProperties.PageNumberAlign.Left
                        cb.AddTemplate(template, pageSize.GetLeft(PageLeftMargin + text.Length + 32), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)))
                    Case PdfGeneratorProperties.PageNumberAlign.Right
                        cb.AddTemplate(template, pageSize.GetRight(PageRightMargin + template.ToString.Length), pageSize.GetBottom(document.PageSize.Bottom + (PageBottomMargin - FooterTable.TotalHeight - 10)))
                End Select

            End If
        End If

        If IsStrokeRequired Then
            cb.MoveTo(PageLeftMargin, pageSize.GetBottom(PageBottomMargin))
            cb.LineTo((document.PageSize.Width - PageRightMargin), pageSize.GetBottom(PageBottomMargin - 2)) 'document.PageSize.Bottom + FooterTable.TotalHeight
            cb.Stroke()
        End If

        If SeprateHeader Then
            document.SetMargins(PageLeftMargin, PageRightMargin, PageTopMargin + ReportTable.TotalHeight(), PageBottomMargin + FooterTable.TotalHeight())
        End If

        myNewPage = False
    End Sub

    Public Overrides Sub OnCloseDocument(ByVal writer As PdfWriter, ByVal document As Document)
        MyBase.OnCloseDocument(writer, document)
        Try
            template.BeginText()
            template.SetFontAndSize(bf, 7.5)
            template.SetTextMatrix(0, 0)

            If TotPagno > pagecount Then
                template.ShowText(Convert.ToString(pagecount))
            Else
                template.ShowText(Convert.ToString((writer.PageNumber - pagecount)))
            End If

            template.EndText()

        Catch ex As Exception

        End Try

    End Sub

    Public Sub NewPage(ByVal document As Document)
        Try
            document.NewPage()
            myNewPage = True
        Catch ex As Exception

        End Try
    End Sub

    Public Sub SetOrientation(ByVal Orientation As PdfNumber)
        Me.orientation = Orientation
    End Sub

End Class

Public Class CustomImageTagProcessor
    Inherits iTextSharp.tool.xml.html.Image
    Public Overrides Function [End](ByVal ctx As IWorkerContext, ByVal tag As Tag, ByVal currentContent As IList(Of IElement)) As IList(Of IElement)
        Dim attributes As IDictionary(Of String, String) = tag.Attributes
        Dim src As String
        If Not attributes.TryGetValue(HTML.Attribute.SRC, src) Then
            Return New List(Of IElement)(1)
        End If

        If String.IsNullOrEmpty(src) Then
            Return New List(Of IElement)(1)
        End If
        If InStr(src, "/") = 0 Then
            src = src & "/"
        End If
        If src.StartsWith("data:image/", StringComparison.InvariantCultureIgnoreCase) Then
            ' data:[<MIME-type>][;charset=<encoding>][;base64],<data>
            Dim base64Data = src.Substring(src.IndexOf(",") + 1)
            Dim imagedata = Convert.FromBase64String(base64Data)
            Dim image = iTextSharp.text.Image.GetInstance(imagedata)

            Dim list = New List(Of IElement)()
            Dim htmlPipelineContext = GetHtmlPipelineContext(ctx)
            list.Add(GetCssAppliers().Apply(New Chunk(DirectCast(GetCssAppliers().Apply(image, tag, htmlPipelineContext), iTextSharp.text.Image), 0, 0, True), tag, htmlPipelineContext))
            Return list
        Else
            Return MyBase.[End](ctx, tag, currentContent)
        End If
    End Function
End Class

Friend Class TableDataProcessor
    Inherits TableData

    Dim wri As PdfWriter
    Sub TableDataProcessor(ByVal writer As PdfWriter)
        wri = writer
    End Sub


    Private Function HasWritingMode(ByVal attributeMap As IDictionary(Of String, String), ByVal Elem As String) As Boolean
        Dim hasStyle As Boolean = attributeMap.ContainsKey("style")
        Return If(hasStyle AndAlso attributeMap("style").Split(New Char() {";"c}).Where(Function(x) x.Contains(Elem)).Count() > 0, True, False)
    End Function

    

    Public Overrides Function [End](ByVal ctx As IWorkerContext, ByVal tag As Tag, ByVal currentContent As IList(Of IElement)) As IList(Of IElement)
        Dim cells = MyBase.[End](ctx, tag, currentContent)
        Dim attributeMap = tag.Attributes
        If HasWritingMode(attributeMap, "writing-mode:") Then
            Dim pdfPCell = DirectCast(cells(0), PdfPCell)
            ' **always** 'sideways-lr'
            pdfPCell.Rotation = 90
        End If

        If HasWritingMode(attributeMap, "white-space:nowrap") Or HasWritingMode(attributeMap, "white-space: nowrap") Then
            For Each cell As IElement In cells
                Dim pdfPCells = DirectCast(cell, PdfPCell)
                pdfPCells.NoWrap = True
            Next
        End If
       

        'If HasWritingMode(attributeMap, "text-transform:uppercase") Or HasWritingMode(attributeMap, "text-transform: uppercase") Then
        '    For Each cell As IElement In cells
        '        Dim pdfPCells = DirectCast(cell, PdfPCell) --works
        '        Dim cc = pdfPCells.Column.CompositeElements

        '        If cc.Count > 0 Then
        '            Dim phase As New Phrase()
        '            For Each chk As iTextSharp.text.ParaGraph In cc
        '                phase.Add(New Phrase(chk.Content.Trim.ToUpper))
        '            Next
        '            pdfPCells.Phrase = phase
        '        End If

        '    Next
        'End If

        Return cells
    End Function
End Class
