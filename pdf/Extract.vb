Imports System.Text
Imports System.IO
Imports iTextSharp.text 'Core PDF Text Functionalities
Imports iTextSharp.text.pdf 'PDF Content
Imports iTextSharp.text.pdf.parser 'Content Parser
Imports pdf.LocTextExtraction
Public Class Extract
    Public Shared Function GetPagesCount(ByVal strSource As String) As Integer
        Dim PgsCnt As Integer = Nothing
        Try
            If File.Exists(strSource) Then 'Check If File Exists
                Using pdfFileReader = New PdfReader(strSource) 'Read Our File
                    PgsCnt = pdfFileReader.NumberOfPages
                End Using
            End If
        Catch ex As System.IO.IOException
            MsgBox(ex.Message)
        End Try
        Return PgsCnt
    End Function
    Public Shared Function FindPDFText(ByVal strSearch As String, _
                                ByVal scCase As StringComparison, _
                                ByVal strSource As String, _
                                ByVal strDest As String, _
                                pbProgress As ProgressBar, _
                                Optional Info As String = "Evry1falls") As Integer
        Dim Matched As Integer = 0
        'Dim psStamp As PdfStamper = Nothing 'PDF Stamper Object
        Dim pcbContent As PdfContentByte = Nothing 'Read PDF Content
        'Source = pdf file
        'strsearch
        If File.Exists(strSource) Then 'Check If File Exists
            Using pdfFileReader = New PdfReader(strSource) 'Read Our File
                Dim NewInfo As Dictionary(Of String, String) = pdfFileReader.Info
                NewInfo("Creator") = Info
                If String.IsNullOrEmpty(strDest) Then
                    MsgBox("Destination file can not be empty path.")
                    Return 0
                    Exit Function
                End If
                Using psStamp = New PdfStamper(pdfFileReader, New FileStream(strDest, FileMode.Create))
                    'Change info
                    psStamp.MoreInfo = NewInfo
                    'Read Underlying Content of PDF File
                    'Loop Through All Pages
                    For intCurrPage As Integer = 1 To pdfFileReader.NumberOfPages
                        'Read PDF File Content Blocks
                        Dim lteStrategy As LocTextExtractionStrategy = _
                            New LocTextExtractionStrategy
                        'Look At Current Block
                        pcbContent = psStamp.GetUnderContent(intCurrPage)
                        'Determine Spacing of Block To See If It Matches Our Search String
                        lteStrategy.UndercontentCharacterSpacing = pcbContent.CharacterSpacing
                        lteStrategy.UndercontentHorizontalScaling = pcbContent.HorizontalScaling
                        'Trigger The Block Reading Process
                        Dim currentText As String = _
                            PdfTextExtractor.GetTextFromPage(pdfFileReader, intCurrPage, lteStrategy)
                        'Determine Match(es)
                        Dim lstMatches As List(Of iTextSharp.text.Rectangle) = _
                            lteStrategy.GetTextLocations(strSearch, scCase)
                        Dim pdLayer As PdfLayer 'Create New Layer
                        'Enable Overwriting Capabilities
                        pdLayer = New PdfLayer("Overrite", psStamp.Writer)
                        'Set Fill Colour Of Replacing Layer
                        pcbContent.SetColorFill(BaseColor.GREEN)
                        'Loop Through Each Match
                        For Each rctRect As Rectangle In lstMatches
                            'Create New Rectangle For Replacing Layer
                            pcbContent.Rectangle(rctRect.Left, rctRect.Bottom, rctRect.Width, rctRect.Height)
                            'Fill With Colour Specified
                            pcbContent.Fill()
                            'Create Layer
                            pcbContent.BeginLayer(pdLayer)
                            'Fill aLyer
                            pcbContent.SetColorFill(BaseColor.BLUE)
                            'Fill Underlying Content
                            pcbContent.Fill()
                            'Create GState Object
                            Dim pgState As PdfGState
                            pgState = New PdfGState()
                            'Set Current State
                            pcbContent.SetGState(pgState)
                            'Fill Letters
                            pcbContent.SetColorFill(BaseColor.GREEN)
                            'Start Text Replace Procedure
                            pcbContent.BeginText()
                            'Get Text Location
                            pcbContent.SetTextMatrix(rctRect.Left, rctRect.Bottom)
                            'Set New Font And Size
                            pcbContent.SetFontAndSize(BaseFont.CreateFont _
                                                      (BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9)
                            'Replacing Text
                            'pcbContent.ShowText("AMAZING!!!!")
                            'Stop Text Replace Procedure
                            pcbContent.EndText()
                            'Stop Layer replace Procedure
                            pcbContent.EndLayer()
                        Next
                        'Set Progressbar Minimum Value
                        pbProgress.Value = 0
                        'Set Progressbar Maximum Value
                        pbProgress.Maximum = lstMatches.Count
                        For I As Integer = 0 To lstMatches.Count - 1
                            'Increase Progressbar Value
                            pbProgress.Value += 1
                            Matched += 1
                        Next
                    Next
                End Using  'Close Stamp Object
            End Using  'Close File
        End If
        'AddPDFWatermark(strSource, strDest, Application.StartupPath & "\watermar2k.png")
        Return Matched
    End Function
    Public Shared Sub AddPDFWatermark(ByVal strSource As String, _
                                      ByVal strDest As String, _
                                      Optional imgSource As String = "watermark.jpg")
        'Dim pdfFileReader As PdfReader = Nothing 'Read File
        'Dim psStamp As PdfStamper = Nothing 'PDF Stamper Object
        Dim imgWaterMark As Image = Nothing 'Watermark Image
        Dim pcbContent As PdfContentByte = Nothing 'Read PDF Content
        Dim rctRect As Rectangle = Nothing 'Create New Rectangle To Host Image
        Dim sngX, sngY As Single 'Page Dimensions
        Dim intPageCount As Integer = 0 'Page Count
        Try
            Using pdfFileReader = New PdfReader(strSource) 'Read File
                rctRect = pdfFileReader.GetPageSizeWithRotation(1) 'Store Page Size
                'Create new Stamper Object
                Using psStamp = New PdfStamper(pdfFileReader, _
                                         New System.IO.FileStream(strDest, System.IO.FileMode.Create))
                    'Get Image To Be Used For The Watermark
                    imgWaterMark = Image.GetInstance(imgSource)
                    'Make Sure Image Can Fit On Page
                    If imgWaterMark.Width > rctRect.Width _
                        OrElse imgWaterMark.Height > rctRect.Height Then
                        imgWaterMark.ScaleToFit(rctRect.Width, rctRect.Height)
                        sngX = (rctRect.Width - imgWaterMark.ScaledWidth) / 2
                        sngY = (rctRect.Height - imgWaterMark.ScaledHeight) / 2
                    Else 'Put In Center Of Page
                        sngX = (rctRect.Width - imgWaterMark.Width) / 2
                        sngY = (rctRect.Height - imgWaterMark.Height) / 2
                    End If
                    imgWaterMark.SetAbsolutePosition(sngX, sngY)
                    'Apply To All Pages-------------------------
                    intPageCount = pdfFileReader.NumberOfPages()
                    For i As Integer = 1 To intPageCount
                        pcbContent = psStamp.GetUnderContent(i)
                        pcbContent.AddImage(imgWaterMark)
                    Next
                    '------------------------------------------
                End Using 'psStamp.Close()
            End Using 'pdfFileReader.Close()
        Catch ex As Exception
            Throw ex 'Something Went Wrong
        End Try
    End Sub
End Class
