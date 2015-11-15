Attribute VB_Name = "LinkToImage"
' Copyright 2015, Martin Dreier
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.

'Constants for messages
Const MSG_NO_HYPERLINKS As Integer = 1
Const MSG_REPLACE_LINK As Integer = 2
Const MSG_REPLACE_TITLE As Integer = 3
Const MSG_ERROR_TITLE As Integer = 4
Const MSG_CANNOT_REPLACE As Integer = 5

Sub LinkToImage()
-Attribute LinkToImage.VB_Description = "Link2Image"
'
' Link2Image Makro
' Ersetzt Hyperlinks im aktiven Dokument durch die Bilder, auf die diese verweisen
'
' Replaces hyperlinks by the images to which they link
'
    Dim hlink As Hyperlink
    Dim msg As String
    Dim choice As VbMsgBoxResult
    Dim index As Integer
    Dim links() As Hyperlink
    Dim currentObject As Variant

'   Exit if no hyperlinks are found in current document
    If ActiveDocument.Hyperlinks.Count < 1 Then
        MsgBox translate(MSG_NO_HYPERLINKS)
        Exit Sub
    End If

'   New array to store hyperlinks, deletion will modify Document.Hyperlinks
'   count and end loop prematurely
    ReDim links(ActiveDocument.Hyperlinks.Count - 1)
    For index = 1 To ActiveDocument.Hyperlinks.Count
        Set links(index - 1) = ActiveDocument.Hyperlinks(index)
    Next
    
'   Loop through all hyperlinks in active document and ask user
'   which links should be replaced
    For Each currentObject In links
        Set hlink = currentObject
        msg = translate(MSG_REPLACE_LINK) & vbCrLf & hlink.Address
        choice = MsgBox(msg, vbYesNoCancel, translate(MSG_REPLACE_TITLE))
        Select Case choice
            Case VbMsgBoxResult.vbYes
                'Replace link with target
                Call replaceHyperlink(hlink)
            Case VbMsgBoxResult.vbNo
                ' Continue with next item
            Case VbMsgBoxResult.vbCancel
                ' Exit processing
                Exit For
        End Select
    Next currentObject
End Sub

' Replace hyperlink. The hyperlink is removed from the document and the target
' image inserted in its place
' hlink: Hyperlink to replace
Function replaceHyperlink(ByVal hlink As Hyperlink)
    On Error GoTo ErrorHandler
    
    Dim target As String
    'Step 1: Save hyperlink target and format for path
    target = hlink.Address
        
    'Step 2: Select hyperlink and delete it
    hlink.Range.Select
    Selection.Delete
    
    'Step 3: Insert picture
    Selection.InlineShapes.AddPicture FileName:=target, LinkToFile:=False, SaveWithDocument:=True
    Exit Function

ErrorHandler:
    MsgBox translate(MSG_CANNOT_REPLACE) & vbCrLf & Err.Description, vbOKOnly, translate(MSG_ERROR_TITLE)
    Resume Next
End Function

' Translate messages into the application language. If the translation for
' the application language is not available, it defaults to English
' messageId: Message ID
' Returns: Translated message, or generic error if message IF is unknown
Function translate(ByVal messageId As Integer) As String
    If Application.Language = msoLanguageIDGerman Then
        Select Case messageId
            Case MSG_NO_HYPERLINKS
                translate = "Das aktuelle Dokument enthält keine Hyperlinks"
            Case MSG_REPLACE_LINK
                translate = "Link ersetzen?"
            Case MSG_REPLACE_TITLE
                translate = "Hyperlink durch Ziel ersetzen"
            Case MSG_ERROR_TITLE
                translate = "Fehler"
            Case MSG_CANNOT_REPLACE
                translate = "Link konnte nicht ersetzt werden. Fehler:"
        End Select
    Else ' Default: English
        Select Case messageId
            Case MSG_NO_HYPERLINKS
                translate = "The active document contains no hyperlinks"
            Case MSG_REPLACE_LINK
                translate = "Replace Link?"
            Case MSG_REPLACE_TITLE
                translate = "Replace Hyperlink With Target"
            Case MSG_ERROR_TITLE
                translate = "Error"
            Case MSG_CANNOT_REPLACE
                translate = "Link could not be replaced. Error:"
        End Select
    End If
    
    If translate = "" Then
        translate = "Unknown message ID: " & messageId
    End If
End Function

