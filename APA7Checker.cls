VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APA7Checker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' APA7Checker Macro
'
'
Sub APA7Checker()
Dim para As Paragraph
Dim refText As String
Dim issues As Integer
Dim pagePattern As String
Dim authorPattern As String
Dim pubYear As Integer
Dim yearString As String

issues = 0
' Regex pattern to find page ranges with incorrect dashes
pagePattern = "\d{1}[-—]\d{1}"
' Regex pattern to match valid author formats, including Dutch last names
authorPattern = "([A-Z][a-z]*,|van \b[A-Z][a-z]*) *"

' Loop through each paragraph
For Each para In ActiveDocument.Paragraphs
    refText = para.Range.Text
    
    ' Check for page range using incorrect dash before "http"
    Dim regexPage As Object
    Set regexPage = CreateObject("VBScript.RegExp")
    regexPage.Pattern = pagePattern
    regexPage.Global = True
    
    ' Check if "http" exists in the text
    If InStr(refText, "http") > 0 Then
        ' Find the position of "http"
        httpPos = InStr(refText, "http")
        ' Extract the text before "http"
        refTextBeforeHttp = Left(refText, httpPos - 1)
        
        ' Check for incorrect dashes in the text before "http"
        If regexPage.test(refTextBeforeHttp) Then
            para.Range.HighlightColorIndex = wdYellow
            para.Range.Comments.Add Range:=para.Range, Text:="Check page range. Use an en-dash (–) instead of a hyphen (-) or em dash (—)."
            issues = issues + 1
        End If
    Else
        ' If "http" is not found, check the entire text
        If regexPage.test(refText) Then
            para.Range.HighlightColorIndex = wdYellow
            para.Range.Comments.Add Range:=para.Range, Text:="Check page range. Use an en-dash (–) instead of a hyphen (-) or em dash (—)."
            issues = issues + 1
        End If
    End If
    
    ' Check for author format only if ".com", ".gov", or ".blog", or "www" is NOT present
    If InStr(refText, ".com") = 0 And InStr(refText, ".gov") = 0 And InStr(refText, ".blog") = 0 And InStr(refText, "www") = 0 Then
        ' Find the position of the year (assuming it's a 4-digit number)
        yearPos = InStrRev(refText, "(", -1) ' Use InStrRev to find the opening bracket before publication year
        
        ' Check if a space was found
        If yearPos > 0 Then
            ' Extract the text before the year
            refTextBeforeYear = Left(refText, yearPos - 1)
            
            ' Check for author format in the text before the year
            Dim regexAuthor As Object
            Set regexAuthor = CreateObject("VBScript.RegExp")
            regexAuthor.Pattern = authorPattern
            regexAuthor.Global = False
            
            If Not regexAuthor.test(refTextBeforeYear) Then
                para.Range.HighlightColorIndex = wdYellow
                para.Range.Comments.Add Range:=para.Range, Text:="Check author formatting. Should start with a valid last name."
                issues = issues + 1
            End If
        End If ' End of If yearPos > 0
        
        ' Check for publication year and DOI
        If yearPos > 0 Then
            ' Extract the potential year string
            yearString = Mid(refText, yearPos + 1, 4)
            
            ' Check if the extracted string is numeric
            If IsNumeric(yearString) Then
                pubYear = CInt(yearString) ' Convert to integer if numeric
                If pubYear >= 2000 Then
                    If InStr(refText, "https://doi.org/") = 0 Then
                        para.Range.HighlightColorIndex = wdYellow
                        para.Range.Comments.Add Range:=para.Range, Text:="Please include DOI."
                        issues = issues + 1
                    End If
                End If
            Else
                ' Show message box if yearString is not numeric
                ' MsgBox "I am not numeric!", vbExclamation
            End If
        End If
    End If
    
Next para

' Notify user of completion
If issues = 0 Then
    MsgBox "All references appear correctly formatted.", vbInformation
Else
    MsgBox issues & " potential formatting issues found.", vbExclamation
End If
End Sub
