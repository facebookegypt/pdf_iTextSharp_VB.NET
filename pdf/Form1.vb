Imports Awesomium
Imports System.Data.OleDb

Public Class Form1
    Private OFD As OpenFileDialog
    Dim DestfileName As String = Nothing
    Dim SrcFile As String = Nothing
    Private SrchStr As String = Nothing
    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        System.IO.File.Delete(DestfileName)
    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then Close()
    End Sub
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        KeyPreview = True
        TxtPdfPath.ReadOnly = True
        DestfileName = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".pdf"
    End Sub
    Private Sub Label2_Click(sender As System.Object, e As System.EventArgs) Handles Label2.Click
        Try
            Using OFD As New OpenFileDialog With {
                .CheckPathExists = True,
                .Filter = ("PDF File Format *pdf|*.pdf"),
                .DefaultExt = "pdf",
                .Multiselect = False,
                .RestoreDirectory = True,
                .InitialDirectory = (Application.StartupPath),
                .SupportMultiDottedExtensions = False}
                If OFD.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    TxtPdfPath.Text = OFD.FileName
                End If
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Label3_Click(sender As System.Object, e As System.EventArgs) Handles Label3.Click
        SrcFile = TxtPdfPath.Text
        SrchStr = TxtPdfSearch.Text
        'CurrentCultureIgnoreCase 'Capital = Small' Letters
        Label5.Text = Extract.FindPDFText(SrchStr, _
                                        StringComparison.CurrentCultureIgnoreCase, _
                                        SrcFile, _
                                        DestfileName, _
                                        ProgressBar1).ToString
        Label8.Text = Extract.GetPagesCount(SrcFile)
        Try
            With WebControl2
                .ViewType = Awesomium.Core.WebViewType.Offscreen
                .Show()
                .Source = (New System.Uri("file:///" & DestfileName))
            End With
        Catch ex As System.IO.IOException
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub WebControl2_DocumentReady(sender As Object, e As Awesomium.Core.DocumentReadyEventArgs) Handles WebControl2.DocumentReady
        If e.ReadyState = Core.DocumentReadyState.Loaded Then
            Try
                System.IO.File.Copy(DestfileName, Application.StartupPath & "\NewPdf.pdf", True)
            Catch ex As IO.IOException
                MsgBox(ex.Message)
            End Try
        End If
    End Sub
End Class
