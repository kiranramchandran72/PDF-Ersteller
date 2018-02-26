Imports System.IO
Imports Datamate.Mediware.PdfGenerator

Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub btnGenPdf_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGenPdf.ServerClick
        Try

            Dim PrintFileStream As StreamReader
            Dim _FileData As String = ""
            Dim FilePath = HttpContext.Current.Server.MapPath("~") & "/" & "html/DSModified.htm" '"html/Content.html DSNew.html DsHtml.html
            Dim fi = New FileInfo(FilePath)
            PrintFileStream = fi.OpenText()
            _FileData = HttpContext.Current.Server.HtmlDecode(PrintFileStream.ReadToEnd())
            PrintFileStream.Close()
            PrintFileStream.Dispose()
          

            Dim _PdfGenerator As New PDFGenerator()
            '_PdfGenerator.SourceHtmlFileName = ""
            _PdfGenerator.SourceHtmlString = _FileData
            _PdfGenerator.WaterMarkContent = "/Img/bgconfidential.png" '/Img/nabh.jpg bgconfidential.png /Img/nabh.jpg
            _PdfGenerator.IsWaterMarkRequired = False
            _PdfGenerator.WaterMarkContentType = PdfGeneratorProperties.WarterMarkType.Image
            _PdfGenerator.IsPageNumberRequired = True
            _PdfGenerator.PageNumberPosition = PdfGeneratorProperties.PageNumberAlign.Right
            _PdfGenerator.IsStrokeRequired = False
            _PdfGenerator.PageBottomMargin = 1
            _PdfGenerator.PageTopMargin = 1
            _PdfGenerator.PageLeftMargin = 1
            _PdfGenerator.PageRightMargin = 1
            _PdfGenerator.DocPageSize = PdfGeneratorProperties.DocumentPageSize.A4
            _PdfGenerator.DirectPrint = True
            '_PdfGenerator.DocPageHeight = 8.63
            ' _PdfGenerator.DocPageWidth = 5.32

            _PdfGenerator.IsLandScape = False


            'Dim res = _PdfGenerator.GeneratePdfDocument(True)


            Dim fileName = "RES_" + DateTime.Now.ToString("dd_MM_yy_HH_mm_ss") + ".pdf"
            _PdfGenerator.GenerateAndDownloadPdf(fileName, False)

            Dim ErrStr = _PdfGenerator.ErrorString

            If ErrStr.Trim.Length > 0 Then
                Throw New Exception(ErrStr)
            End If

            'For i As Integer = 0 To 1
            '    _PdfGenerator.GenerateAndDownloadPdf(fileName, False)
            'Next



            'Response.ContentType = "application/pdf"
            'Response.AddHeader("content-disposition", "inline;filename=" & fileName)
            'Response.Cache.SetCacheability(HttpCacheability.NoCache)

            'Response.Write(_PdfGenerator.GenerateMultiplePdfDocument(True))
            ''Response.TransmitFile(fileName)
            'HttpContext.Current.ApplicationInstance.CompleteRequest()
        Catch ex As Exception
            AddToLogFile(ex.Message)
        End Try
    End Sub

    Public Function AddToLogFile(ByVal strData As String, _
    Optional ByVal FullPath As String = "", _
      Optional ByVal ErrInfo As String = "") As Boolean

        Dim bAns As Boolean = False
        Dim objWriter As IO.StreamWriter
        Try
            If FullPath.Length = 0 Then
                FullPath = System.Web.HttpContext.Current.Server.MapPath("~") & "/ErrorLog/Errorlog.txt"
            End If
            objWriter = New IO.StreamWriter(FullPath, True)
            objWriter.WriteLine(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt", Globalization.CultureInfo.InvariantCulture) & "    :   " & strData)
            objWriter.Close()

            bAns = True
        Catch Ex As Exception
            ErrInfo = Ex.Message

        End Try

        Try
            If (Not System.Diagnostics.EventLog.SourceExists("EMR")) Then
                System.Diagnostics.EventLog.CreateEventSource("EMR", "Mediware")
                System.Diagnostics.EventLog.WriteEntry("EMR", strData)
            Else
                System.Diagnostics.EventLog.WriteEntry("EMR", strData)
            End If
        Catch ex As Exception

        End Try

        Return bAns
    End Function

End Class
