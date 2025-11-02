Sub FixingEncodingIssuesAdvanced()
    Dim cell As Range
    Dim replacements As Object
    Dim key As Variant
    Dim originalValue As String
    Dim newValue As String
    Dim changedCount As Long
    Dim logFile As Integer
    Dim logPath As String
    Dim timestamp As String
    Dim cellAddress As String
    
    ' Create timestamp for log file
    timestamp = Format(Now(), "yyyy-mm-dd_hh-nn-ss")
    logPath = ThisWorkbook.Path & "\EncodingFixes_" & timestamp & ".txt"
    
    ' Open log file for writing
    logFile = FreeFile
    Open logPath For Output As #logFile
    
    ' Write header to log
    Print #logFile, "=================================="
    Print #logFile, "Encoding Fix Log"
    Print #logFile, "Date: " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
    Print #logFile, "Workbook: " & ThisWorkbook.Name
    Print #logFile, "Worksheet: " & ActiveSheet.Name
    Print #logFile, "=================================="
    Print #logFile, ""
    
    ' Create dictionary of replacements
    Set replacements = CreateObject("Scripting.Dictionary")
    
    ' ====================================================================
    ' UTF-8 TO CP-1252/LATIN-1 DOUBLE-ENCODING ISSUES
    ' ====================================================================
    
    ' Lowercase accented letters (À-ÿ range)
    replacements.Add "Ã ", "à"
    replacements.Add "Ã¡", "á"
    replacements.Add "Ã¢", "â"
    replacements.Add "Ã£", "ã"
    replacements.Add "Ã¤", "ä"
    replacements.Add "Ã¥", "å"
    replacements.Add "Ã¦", "æ"
    replacements.Add "Ã§", "ç"
    replacements.Add "Ã¨", "è"
    replacements.Add "Ã©", "é"
    replacements.Add "Ãª", "ê"
    replacements.Add "Ã«", "ë"
    replacements.Add "Ã¬", "ì"
    replacements.Add "Ã­", "í"
    replacements.Add "Ã®", "î"
    replacements.Add "Ã¯", "ï"
    replacements.Add "Ã°", "ð"
    replacements.Add "Ã±", "ñ"
    replacements.Add "Ã²", "ò"
    replacements.Add "Ã³", "ó"
    replacements.Add "Ã´", "ô"
    replacements.Add "Ãµ", "õ"
    replacements.Add "Ã¶", "ö"
    replacements.Add "Ã·", "÷"
    replacements.Add "Ã¸", "ø"
    replacements.Add "Ã¹", "ù"
    replacements.Add "Ãº", "ú"
    replacements.Add "Ã»", "û"
    replacements.Add "Ã¼", "ü"
    replacements.Add "Ã½", "ý"
    replacements.Add "Ã¾", "þ"
    replacements.Add "Ã¿", "ÿ"
    replacements.Add "ÃŸ", "ß"
    
    ' Uppercase accented letters (À-Ý range)
    replacements.Add "Ã€", "À"
    replacements.Add "Ã", "Á"
    replacements.Add "Ã‚", "Â"
    replacements.Add "Ãƒ", "Ã"
    replacements.Add "Ã„", "Ä"
    replacements.Add "Ã…", "Å"
    replacements.Add "Ã†", "Æ"
    replacements.Add "Ã‡", "Ç"
    replacements.Add "Ãˆ", "È"
    replacements.Add "Ã‰", "É"
    replacements.Add "ÃŠ", "Ê"
    replacements.Add "Ã‹", "Ë"
    replacements.Add "ÃŒ", "Ì"
    replacements.Add "Ã", "Í"
    replacements.Add "ÃŽ", "Î"
    replacements.Add "Ã", "Ï"
    replacements.Add "Ã", "Ð"
    replacements.Add "Ã'", "Ñ"
    replacements.Add "Ã'", "Ò"
    replacements.Add "Ã"", "Ó"
    replacements.Add "Ã"", "Ô"
    replacements.Add "Ã•", "Õ"
    replacements.Add "Ã–", "Ö"
    replacements.Add "Ã—", "×"
    replacements.Add "Ã˜", "Ø"
    replacements.Add "Ã™", "Ù"
    replacements.Add "Ãš", "Ú"
    replacements.Add "Ã›", "Û"
    replacements.Add "Ãœ", "Ü"
    replacements.Add "Ã", "Ý"
    replacements.Add "Ãž", "Þ"
    
    ' ====================================================================
    ' POLISH SPECIAL CHARACTERS
    ' ====================================================================
    replacements.Add "Ä…", "ą"
    replacements.Add "Ä„", "Ą"
    replacements.Add "Ä‡", "ć"
    replacements.Add "Ä†", "Ć"
    replacements.Add "Ä™", "ę"
    replacements.Add "Ä˜", "Ę"
    replacements.Add "Å‚", "ł"
    replacements.Add "Å", "Ł"
    replacements.Add "Å„", "ń"
    replacements.Add "Åƒ", "Ń"
    replacements.Add "Å›", "ś"
    replacements.Add "Åš", "Ś"
    replacements.Add "Åº", "ź"
    replacements.Add "Å¹", "Ź"
    replacements.Add "Å¼", "ż"
    replacements.Add "Å»", "Ż"
    
    ' ====================================================================
    ' CZECH/SLOVAK SPECIAL CHARACTERS
    ' ====================================================================
    replacements.Add "Ä", "č"
    replacements.Add "ÄŒ", "Č"
    replacements.Add "Ä", "ď"
    replacements.Add "ÄŽ", "Ď"
    replacements.Add "Ä›", "ě"
    replacements.Add "Äš", "Ě"
    replacements.Add "Ň", "ň"
    replacements.Add "Å‡", "Ň"
    replacements.Add "Å™", "ř"
    replacements.Add "Å˜", "Ř"
    replacements.Add "Å¡", "š"
    replacements.Add "Å ", "Š"
    replacements.Add "Å¥", "ť"
    replacements.Add "Å¤", "Ť"
    replacements.Add "Å¯", "ů"
    replacements.Add "Å®", "Ů"
    replacements.Add "Å¾", "ž"
    replacements.Add "Å½", "Ž"
    replacements.Add "Å•", "ő"
    replacements.Add "Å"", "Ő"
    replacements.Add "Å±", "ű"
    replacements.Add "Å°", "Ű"
    
    ' ====================================================================
    ' TURKISH SPECIAL CHARACTERS
    ' ====================================================================
    replacements.Add "Ä±", "ı"
    replacements.Add "Ä°", "İ"
    replacements.Add "ÄŸ", "ğ"
    replacements.Add "Äž", "Ğ"
    replacements.Add "ÅŸ", "ş"
    replacements.Add "Åž", "Ş"
    
    ' ====================================================================
    ' ROMANIAN SPECIAL CHARACTERS
    ' ====================================================================
    replacements.Add "Å£", "ţ"
    replacements.Add "Å¢", "Ţ"
    replacements.Add "È™", "ș"
    replacements.Add "È˜", "Ș"
    replacements.Add "È›", "ț"
    replacements.Add "Èš", "Ț"
    
    ' ====================================================================
    ' LATVIAN/LITHUANIAN SPECIAL CHARACTERS
    ' ====================================================================
    replacements.Add "Ä", "ā"
    replacements.Add "Ä€", "Ā"
    replacements.Add "Ä"", "ē"
    replacements.Add "Ä'", "Ē"
    replacements.Add "Ä£", "ģ"
    replacements.Add "Ä¢", "Ģ"
    replacements.Add "Ä«", "ī"
    replacements.Add "Äª", "Ī"
    replacements.Add "Ä·", "ķ"
    replacements.Add "Ä¶", "Ķ"
    replacements.Add "Ä¼", "ļ"
    replacements.Add "Ä»", "Ļ"
    replacements.Add "Å†", "ņ"
    replacements.Add "Å…", "Ņ"
    replacements.Add "Å«", "ū"
    replacements.Add "Åª", "Ū"
    
    ' ====================================================================
    ' ICELANDIC/NORDIC SPECIAL CHARACTERS
    ' ====================================================================
    replacements.Add "Ã°", "ð"
    replacements.Add "Ã", "Ð"
    replacements.Add "Ã¾", "þ"
    replacements.Add "Ãž", "Þ"
    
    ' ====================================================================
    ' CP-1252 SPECIFIC: SMART QUOTES AND PUNCTUATION (0x80-0x9F range)
    ' ====================================================================
    ' These are the special CP-1252 characters that don't exist in ISO-8859-1
    
    ' Quotes
    replacements.Add "â€œ", """    ' Left double quote (U+201C)
    replacements.Add "â€", """     ' Right double quote (U+201D)
    replacements.Add "â€˜", "'"     ' Left single quote (U+2018)
    replacements.Add "â€™", "'"     ' Right single quote (U+2019)
    replacements.Add "â€ž", "„"     ' Double low quote (U+201E)
    replacements.Add "â€º", "›"     ' Single right angle quote (U+203A)
    replacements.Add "â€¹", "‹"     ' Single left angle quote (U+2039)
    replacements.Add "Â«", "«"      ' Left guillemet
    replacements.Add "Â»", "»"      ' Right guillemet
    
    ' Dashes
    replacements.Add "â€"", "–"     ' En dash (U+2013)
    replacements.Add "â€"", "—"     ' Em dash (U+2014)
    replacements.Add "â€•", "―"     ' Horizontal bar (U+2015)
    
    ' Other punctuation
    replacements.Add "â€¦", "…"     ' Ellipsis (U+2026)
    replacements.Add "â€¢", "•"     ' Bullet (U+2022)
    replacements.Add "â€°", "‰"     ' Per mille (U+2030)
    replacements.Add "â€ ", "†"     ' Dagger (U+2020)
    replacements.Add "â€¡", "‡"     ' Double dagger (U+2021)
    
    ' ====================================================================
    ' CURRENCY AND SYMBOLS
    ' ====================================================================
    replacements.Add "â‚¬", "€"     ' Euro
    replacements.Add "Â£", "£"      ' Pound
    replacements.Add "Â¥", "¥"      ' Yen
    replacements.Add "Â¢", "¢"      ' Cent
    replacements.Add "Â¤", "¤"      ' Currency sign
    
    ' Copyright and trademark
    replacements.Add "Â©", "©"      ' Copyright
    replacements.Add "Â®", "®"      ' Registered
    replacements.Add "â„¢", "™"     ' Trademark
    
    ' Math and special symbols
    replacements.Add "Â°", "°"      ' Degree
    replacements.Add "Â±", "±"      ' Plus-minus
    replacements.Add "Â²", "²"      ' Superscript 2
    replacements.Add "Â³", "³"      ' Superscript 3
    replacements.Add "Â¹", "¹"      ' Superscript 1
    replacements.Add "Âµ", "µ"      ' Micro
    replacements.Add "Â¶", "¶"      ' Pilcrow (paragraph)
    replacements.Add "Â·", "·"      ' Middle dot
    replacements.Add "Â¸", "¸"      ' Cedilla
    replacements.Add "Âº", "º"      ' Masculine ordinal
    replacements.Add "Âª", "ª"      ' Feminine ordinal
    replacements.Add "Â´", "´"      ' Acute accent
    replacements.Add "Â¨", "¨"      ' Diaeresis
    replacements.Add "Â¯", "¯"      ' Macron
    replacements.Add "Â¬", "¬"      ' Not sign
    replacements.Add "Â­", "­"      ' Soft hyphen
    
    ' Fractions
    replacements.Add "Â¼", "¼"      ' 1/4
    replacements.Add "Â½", "½"      ' 1/2
    replacements.Add "Â¾", "¾"      ' 3/4
    
    ' Punctuation
    replacements.Add "Â¿", "¿"      ' Inverted question mark
    replacements.Add "Â¡", "¡"      ' Inverted exclamation
    
    ' ====================================================================
    ' HTML ENTITIES (in case they appear in data)
    ' ====================================================================
    replacements.Add "&#39;", "'"
    replacements.Add "&quot;", """"
    replacements.Add "&amp;", "&"
    replacements.Add "&lt;", "<"
    replacements.Add "&gt;", ">"
    replacements.Add "&nbsp;", " "
    replacements.Add "O&#39;", "O'"
    
    ' ====================================================================
    ' COMMON CORRUPTION PATTERNS
    ' ====================================================================
    replacements.Add "ï¿½", "�"     ' Unicode replacement character
    replacements.Add "Ð", "'"       ' Often a corrupted apostrophe
    replacements.Add "â€�", "-"     ' Corrupted dash
    replacements.Add "Ã‚", ""       ' Often appears as phantom character
    
    ' ====================================================================
    ' TRIPLE-ENCODED ISSUES (UTF-8 encoded twice)
    ' ====================================================================
    ' These are less common but can happen with multiple import/export cycles
    replacements.Add "ÃƒÂ©", "é"
    replacements.Add "ÃƒÂ¨", "è"
    replacements.Add "ÃƒÂª", "ê"
    replacements.Add "ÃƒÂ«", "ë"
    replacements.Add "ÃƒÂ ", "à"
    replacements.Add "ÃƒÂ¢", "â"
    replacements.Add "ÃƒÂ´", "ô"
    replacements.Add "ÃƒÂ»", "û"
    replacements.Add "ÃƒÂ¼", "ü"
    replacements.Add "ÃƒÂ¶", "ö"
    replacements.Add "ÃƒÂ¤", "ä"
    replacements.Add "ÃƒÂ§", "ç"
    replacements.Add "ÃƒÂ±", "ñ"
    replacements.Add "ÃƒÅ¡", "š"
    replacements.Add "ÃƒÅ½", "Ž"
    
    ' ====================================================================
    ' PROCESS CELLS
    ' ====================================================================
    
    changedCount = 0
    Application.ScreenUpdating = False  ' Speed up processing
    
    For Each cell In ActiveSheet.UsedRange.Cells
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            originalValue = cell.Value
            newValue = originalValue
            
            ' Apply all replacements
            For Each key In replacements.Keys
                If InStr(newValue, key) > 0 Then
                    newValue = Replace(newValue, key, replacements(key))
                End If
            Next key
            
            ' Only update if changed
            If newValue <> originalValue Then
                cell.Value = newValue
                changedCount = changedCount + 1
                cellAddress = cell.Address(False, False)
                
                ' Write to log file
                Print #logFile, "Row " & cell.Row & ", Column " & cell.Column & " (" & cellAddress & "):"
                Print #logFile, "  BEFORE: " & originalValue
                Print #logFile, "  AFTER:  " & newValue
                Print #logFile, ""
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    ' Write summary to log
    Print #logFile, "=================================="
    Print #logFile, "SUMMARY"
    Print #logFile, "=================================="
    Print #logFile, "Total cells processed: " & ActiveSheet.UsedRange.Cells.Count
    Print #logFile, "Cells with fixes: " & changedCount
    Print #logFile, "Encoding patterns checked: " & replacements.Count
    Print #logFile, ""
    Print #logFile, "Log file saved to: " & logPath
    
    ' Close log file
    Close #logFile
    
    ' Notify user
    If changedCount > 0 Then
        MsgBox "Fixed " & changedCount & " cell(s) with encoding issues!" & vbCrLf & vbCrLf & _
               "Log file saved to:" & vbCrLf & logPath & vbCrLf & vbCrLf & _
               "Party on Wayne!", vbInformation, "Encoding Fixes Complete"
    Else
        MsgBox "No encoding issues found. Your data is clean!" & vbCrLf & vbCrLf & _
               "Log file saved to:" & vbCrLf & logPath, vbInformation, "Encoding Check Complete"
    End If
    
End Sub
