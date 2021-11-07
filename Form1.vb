Imports System.IO
Imports System.Net.Mail
Imports System.Text

Public Class Form1
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        With Me.HtmlEditControl1
            .Dock = DockStyle.Fill
            .ShowPropertyGrid = False
            .CSSText = "BODY {font-family: Arial Unicode MS, Arial, Sans-Serif;}"
            .LicenceKey = "YourLicenseKey"
            .LicenceKeyInlineSpelling = "YourInlineSpellingLicenseKey"
            .EnableInlineSpelling = True
            .BaseURL = Application.StartupPath.Replace("\", "/")
            .UseRelativeURLs = True
            .ImageStorageLocation = Path.Combine(Application.StartupPath, "image")
        End With

        Dim oButton = Me.HtmlEditControl1.ToolStripItems.Add("Save As Email")
        AddHandler oButton.Click, AddressOf AsEmail_Click
        oButton.Padding = New Padding(3)
        Dim oButtonSpell = Me.HtmlEditControl1.ToolStripItems.Add("Check Spelling")
        AddHandler oButtonSpell.Click, AddressOf OButtonSpell_Click
        oButtonSpell.Padding = New Padding(3)
        AddHandler Me.HtmlEditControl1.SpellCheckComplete, AddressOf OEdit_SpellCheckComplete

    End Sub

    Private Sub OEdit_SpellCheckComplete(sender As Object, e As EventArgs)
        MessageBox.Show("Spelling Check Completed")
    End Sub

    Private Sub OButtonSpell_Click(sender As Object, e As EventArgs)
        Me.HtmlEditControl1.SpellCheckDocument(True)
    End Sub

    Private Sub AsEmail_Click(sender As Object, e As EventArgs)

        Dim EditorHTML As String = Me.HtmlEditControl1.DocumentHTML
        Dim newMail As MailMessage = New MailMessage()
        newMail.[To].Add(New MailAddress("you@your.address"))
        newMail.From = (New MailAddress("someone@their.address"))
        newMail.Subject = "Test Subject"
        newMail.IsBodyHtml = True
        Dim inlineLogoList As List(Of LinkedResource) = New List(Of LinkedResource)()

        For Each oImage As HtmlElement In Me.HtmlEditControl1.GetItemsByTagName("img")
            Dim oUri As New Uri(oImage.GetAttribute("src"))

            If oUri.IsFile Then
                Dim inlineLogo = New LinkedResource(oUri.LocalPath, "image/" & New FileInfo(oUri.LocalPath).Extension.Substring(1)) With {
                .ContentId = Guid.NewGuid().ToString()
            }
                oImage.SetAttribute("src", "cid:" & inlineLogo.ContentId)
                inlineLogoList.Add(inlineLogo)
            End If
        Next

        Dim view = AlternateView.CreateAlternateViewFromString(Me.HtmlEditControl1.Document.Body.OuterHtml, Nothing, "text/html")

        For Each inlineLogo As LinkedResource In inlineLogoList
            view.LinkedResources.Add(inlineLogo)
        Next

        newMail.AlternateViews.Add(view)
        Dim view2 = AlternateView.CreateAlternateViewFromString(Me.HtmlEditControl1.Document.Body.InnerText & vbNullString, Encoding.ASCII, "text/plain")
        newMail.AlternateViews.Add(view2)
        Dim client As SmtpClient = New SmtpClient("mysmtphost") With {
        .DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory,
        .PickupDirectoryLocation = Application.StartupPath & "/SMTPPickupFolder"
    }
        client.Send(newMail)
        Me.HtmlEditControl1.DocumentHTML = EditorHTML
        MessageBox.Show("Successfully saved to " & vbCrLf & vbCrLf & " " & Application.StartupPath & "/SMTPPickupFolder")
    End Sub
End Class
