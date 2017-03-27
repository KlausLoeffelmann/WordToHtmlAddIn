Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Tools.Ribbon

Public Class HtmlExportRibbon

    Private Sub HtmlExportButtonClick_Event(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        Dim activeRange As Range = ThisAddIn.AddInReference?.Application?.ActiveDocument?.Range

        If activeRange Is Nothing Then
            MessageBox.Show("Sorry, there is no text to export!",
                            "Can't export.", MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        Dim frmProgress As New ProgressForm
        frmProgress.Show()
        frmProgress.MaxChars = activeRange.Text.Length

        System.Windows.Forms.Application.DoEvents()

        Dim sb As New StringBuilder
        Dim charCount = 0

        Dim sw = Stopwatch.StartNew
        Dim currentParagraphStyle As String = Nothing
        Dim currentHyperlinkObject As Hyperlink = Nothing
        Dim isBold = False
        Dim isItalic = False
        Dim isParagraphStarted = False
        Dim isBullet = False
        Dim isNumBullet = False
        Dim crLfCounter = 0
        Dim crLfFound = True
        Dim listInProgress = False
        Dim listingInProgress = False

        Dim currentHyperlinkText As String = Nothing
        Dim currentHyperlink As Hyperlink = Nothing

        For Each currChar As Range In activeRange.Characters

            If currChar.Bold = -1 And Not isBold Then
                sb.Append("<strong>")
                isBold = True
            End If
            If currChar.Italic = -1 And Not isItalic Then
                sb.Append("<i>")
                isItalic = True
            End If
            If currChar.Italic = 0 And isItalic Then
                sb.Append("</i>")
                isItalic = False
            End If
            If currChar.Bold = 0 And isBold Then
                sb.Append("</strong>")
                isBold = False
            End If

            If currChar.Text = vbCr Then
                crLfFound = True
                Continue For
            End If

            If crLfFound Then
                crLfFound = False
                Dim tmpCurrCharParagraphStyle = CType(currChar.ParagraphStyle, Style).NameLocal
                If currentParagraphStyle = tmpCurrCharParagraphStyle Then
                    If listInProgress Then
                        sb.Append("</li>" & vbCrLf & "<li>")
                    ElseIf listingInProgress Then
                        sb.Append(vbCrLf)
                    End If

                    If currentParagraphStyle = "Standard" Then
                        sb.Append("</p>" & vbCrLf & "<p>")
                    End If
                Else
                    If currentParagraphStyle = "Titel" Then
                        sb.Append("</titel>" & vbCrLf)
                    End If
                    If currentParagraphStyle = "Standard" Then
                        sb.Append("</p>" & vbCrLf)
                    End If
                    If currentParagraphStyle = "Überschrift 1" Or currentParagraphStyle = "Heading 1" Then
                        sb.Append("</h1>" & vbCrLf)
                    End If
                    If currentParagraphStyle = "Überschrift 2" Or currentParagraphStyle = "Heading 2" Then
                        sb.Append("</h2>" & vbCrLf)
                    End If
                    If currentParagraphStyle = "Überschrift 3" Or currentParagraphStyle = "Heading 3" Then
                        sb.Append("</h3>" & vbCrLf)
                    End If
                    If currentParagraphStyle = "Listenabsatz" Or currentParagraphStyle = "?" Then
                        sb.Append("</li>" & vbCrLf & "</ul>" & vbCrLf)
                        listInProgress = False
                    End If


                    If tmpCurrCharParagraphStyle = "Titel" Then
                        sb.Append("<titel>")
                    End If
                    If tmpCurrCharParagraphStyle = "Standard" Then
                        sb.Append("<p>")
                    End If
                    If tmpCurrCharParagraphStyle = "Überschrift 1" Or tmpCurrCharParagraphStyle = "Heading 1" Then
                        sb.Append("<h1>")
                    End If
                    If tmpCurrCharParagraphStyle = "Überschrift 2" Or tmpCurrCharParagraphStyle = "Heading 2" Then
                        sb.Append("<h2>")
                    End If
                    If tmpCurrCharParagraphStyle = "Überschrift 3" Or tmpCurrCharParagraphStyle = "Heading 3" Then
                        sb.Append("<h3>")
                    End If
                    If tmpCurrCharParagraphStyle = "Listenabsatz" Or tmpCurrCharParagraphStyle = "?" Then
                        sb.Append("<ul>" & vbCrLf & "<li>")
                        listInProgress = True
                    End If
                    currentParagraphStyle = tmpCurrCharParagraphStyle
                End If
            End If

            If currChar.Hyperlinks?.Count > 0 And currentHyperlink Is Nothing Then
                'Start collecting Hyperlink Display Text.
                currentHyperlinkText = currChar.Text
                'It's VB6, so it is 1 based! :-)
                currentHyperlink = DirectCast(currChar.Hyperlinks(1), Hyperlink)
                Continue For
            End If

            If currChar.Hyperlinks?.Count > 0 And currentHyperlink IsNot Nothing Then
                currentHyperlinkText &= currChar.Text
                Continue For
            End If

            If currChar.Hyperlinks?.Count = 0 And currentHyperlink IsNot Nothing Then
                sb.Append("<a href=""" & currentHyperlink.Address & """ target=""""_blank"""">" & currentHyperlinkText & "</a>")
                currentHyperlink = Nothing
            End If

            sb.Append(currChar.Text)

            charCount += 1
            If charCount Mod 10 = 0 Then
                frmProgress.Progress = charCount
            End If

        Next

        sw.Stop()

        frmProgress.Hide()
        Dim frmResultForm = New ResultForm

        frmResultForm.ShowDialog(sb.ToString)

    End Sub
End Class

Public Class StyleTagTranslateItem
    Property GermanStyleName As String
    Property EnglishStyleName As String
    Property StartTag As String
    Property EndTag As String
    Property IsCharacterStyle As Boolean

End Class
