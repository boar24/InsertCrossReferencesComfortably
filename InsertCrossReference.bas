Attribute VB_Name = "InsertCrossReference"
Option Explicit

Sub InsertCrossReference_()
    ' Macro to insert cross references comfortably
    '
    ' Preparation:
    ' 1) Make sure, the following References are ticked in the VBA editor:
    '       - Microsoft VBScript Regular Expressions 5.5
    '    How to do it: https://www.datanumen.com/blogs/add-object-library-reference-vba/
    ' 2) Put this macro code in a VBA module in your document or document template.
    '    It is recommended to put it into *normal.dot*,
    '    then the functionality is available in any document.
    ' 3) Assign a keyboard shortcut to this macro (recommendation: Ctrl+Alt+Q)
    '    This works like this (in Office 2010):
    '      - File -> Options -> Adapt Ribbon -> Keyboard Shortcuts: Modify...
    '      - Select from Categories: Macros
    '      - Select form Macros: [name of the Macro]
    '      - Assign keyboard shortcut...
    ' 4) Alternatively to 3) or in addition to the shortcut, you can assign this
    '    macro to the ribbon button *Insert -> CrossReference*.
    '    However, then you will not be able any more to access Word's dialog
    '    for inserting cross references.
    '      To assign this sub to the ribbon button *Insert -> CrossReference*,
    '      just rename this sub to *InsertCrossReference* (without underscore).
    '      To de-assign, re-rename it to something like
    '      *InsertCrossReference_* (with underscore).
    '
    ' Useage:
    ' 1) Put the cursor to the location in the document, where the
    '    crossreference shall be inserted,
    '    then press the keyboard shortcut.
    '    A temporary bookmark is inserted
    '    (if their display is enabled, grey square brackets will appear).
    ' 2) Move the cursor to the location to where the crossref shall point.
    '    Supported are:
    '    * Bookmarks
    '        e.g. "[42]" (bibliographic reference)
    '    * Headlines
    '    * Subtitles of Figures, Abbildungen, Tables, Tabellen, etc.
    '        e.g. "Figure 123", "Figure 12-345"
    '    * Subtitles of Tables  realised via { SEQ Table} ,
    '        e.g. "Table 123", "Table 12-345"
    '    * More of the above can be configured, see below.
    '    Hint: Recommendation for large documents: use the navigation pane
    '          (View -> Navigation -> Headlines) to quickly find the
    '          destination location.
    '    Hint: Cross references to hidden text are not possible
    '    Hint: The macro may fail trying to cross reference to locations that
    '          have heavily been edited (deletions / moves) while
    '          "track changes" (markup mode) was active.
    ' 3) Press the keyboard shortcut again.
    '    The cursor will jump back to the location of insertion
    '    and the crossref will be inserted. Done!
    ' 4) Additional function:
    '    Positon the cursor at a cross reference field
    '    (if you have configured chained cross reference fields,
    '    put the cursor to the last field in the chain).
    '    Press the keyboard shortcut.
    '    - The field display toggles to the next configured option,
    '      e.g. from see Chapter 1 to cf. Introduction.
    '    - Subsequently added cross references will use the latest format
    '      (persistent until Word is exited).
    '
    '    You can configure multiple options on how the cross references
    '    shall be inserted,
    '       e.g. as "Figure 12" or "Figure 12 - Overview" etc..
    '    See below under "=== Configuration" on how
    '    to modify the default configuration.
    '    Once configured, you can toggle between the different options
    '    one after the other as follows:
    '    - put the cursor inside a cross reference field, or immediately
    '      behind it (it is generally recommended to set <Field shading>
    '      to <Always> (see https://wordribbon.tips.net/T006107_Controlling_Field_Shading.html)
    '      in order to have fields highlighted by grey background.
    '    - press the keyboard shortcut
    '    - that field toggles its display to be according to the next option
    '    - the current selection is remembered for subsequent
    '      cross reference inserts (persistent until closure of Word)
    '
    ' Revision History:
    ' 151204 Beginn der Revision History
    ' 160111 Kann jetzt auch umgehen mit Numerierungen mit Bindestrich à la "Figure 1-1"
    ' 160112 Jetzt auch Querverweise möglich auf Dokumentenreferenzen à la "[66]" mit Feld " SEQ Ref "
    ' 160615 Felder werden upgedatet falls nötig
    ' 180710 Support für "Nummeriertes Element"
    ' 181026 Generischerer Code für Figure|Table|Abbildung
    ' 190628 New function: toggle to insert numeric or text references ("\r")
    ' 190629 Explanations and UI changed to English
    ' 190705 More complete and better configurable inserts
    ' 190709 Expanded configuration possibilities due to text sequences

    Static isActive As Boolean                  ' remember whether we are in insertion mode
    Static cfgPHeadline As Integer              ' ptr to current config for Headlines
    Static cfgPBookmark As Integer              ' ptr to current config for Bookmarks
    Static cfgPFigureTE As Integer              ' ptr to current config for Figures, Tables, ...

    Dim cfgHeadline As String                   ' configurations for Headlines
    Dim cfgBookmark As String                   ' configurations for Bookmarks
    Dim cfgFigureTE As String                   ' configurations for Figures, Tables, ...
    Dim cfgAHeadline() As String                ' Array with configs for Headlines
    Dim cfgABookmark() As String                ' Array with configs for Bookmarks
    Dim cfgAFigureTE() As String                ' Array with configs for Figures, Tables, ...
    
    Dim paramRefType As Variant                 ' type of reference (WdReferenceType)
    Dim paramRefKind As Variant                 ' kind of reference (WdReferenceKind)
    Dim paramRefText As Variant                 ' content of the field
    Dim paramRefReal As String                  ' which of the three configurations
    
    Dim Response, storeTrackStatus, lastpos As Variant
    Dim prompt As String
    Dim retry, found As Boolean
    Dim index  As Variant
    Dim myerrtxt As String
    Dim linktype, searchstring As String
    Dim linktypes() As String
    Dim allowed As Boolean


    ' ============================================================================================
    ' === Configuration
    ' The following defines how the crossreferences are inserted.
    ' You may reconfigure according to your preferences. Or just use the defined defaults.
    '
    ' For a basic understanding, it is helpful to know the hierarchy of configurations:
    '   configurations
    '       options
    '           parts
    '               switches
    '
    ' There are three configurations according to the three types of fields:
    '   cfgHeadline for Headlines
    '   cfgBookmark for Bookmark
    '   cfgFigureTE for Figures/Tables/Equations/...
    '
    ' For each configuration, multiple options can be configured.
    ' These are the options between which you can toggle. Accordingly, a certain reference
    ' would be displayed e.g as
    '    Figure 1 - System overview
    '    Figure 1
    '    System overview
    '    ...
    '
    ' Each option can have multiple parts, where parts are
    '   either Field code sequences          (example: <REF \h \r>)
    '   or     text       sequences          (example: <see chapter >).
    ' Text sequences must be enclosed in <'> (example: <' - '>).
    ' The <£> is used to represent a non-breaking space.
    '
    ' Each Field code sequence can have multiple switches (example: <\h \r>).
    '
    '
    ' When Word is started, the configurations always default to the first option.
    ' When you toggle, you switch to the next option of the configuration. After the last defined
    ' option, the first reappears. Toggling applies to the selected type of fields only,
    ' thus there are three independent toggles for Headlines, Bookmarks and FigureTEs.
    ' After toggling, the selected option is remembered for subsequent inserts of the
    ' respective type. The memory is persistent until Word is closed.
    ' The individual options are defined in the configuration string, one after the other,
    ' seperated by the sign <|>.
    '
    ' The meaning of the individual switches is similar to the switches of Word's {REF} field:
    '   Main switches:
    '     <REF>,<R>     element's name
    '     <PAGEREF>,<P> insert pagenumber instead of reference
    '     <' '>         text sequence, allows to combine cross references to things like
    '                   <(see chapter 32 on page 48 BELOW)>
    '   Modifier switches:
    '     <\r>          Number instead of text
    '     <\p>          insert <above> or <below> (or whatever it is in your local language)
    '     <\n>          no context                              (not applicable to cfgFigure)
    '     <\w>          full context                            (not applicable to cfgFigure)
    '     <\c>          combination of category + number + text (not applicable to cfgBookmark)
    '     <\h>          insert the cross reference as a hyperlink
    '
    '
    ' Configuration for Headlines:
    cfgHeadline = "R \r  |REF |R \r '£-£'R  |'(see chapter 'R \r' on page 'PAGEREF')'|R \r ' on p.£'PAGEREF|R \p       "
    '             "number|text|number°-°text| (see chapter  XX    on page YY       ) |number on p.°XX      |above/below"
    '
    ' Configuration for Bookmarks:
    cfgBookmark = "R    |PAGEREF|R \p       |R  ' (see£' R \p    ')'"
    '             "text |pagenr |above/below|text (see°above/below) "
    '
    ' Configuration for Figures, Tables, Equations, ...:
    cfgFigureTE = "R \r     |R \r    '£-£'R  |R   |P     |R \p       |R \c            "
    '             "Figure xx|Figure xx - desc|desc|pagenr|above/below|Figure xxTabdesc"
    
    ' Favourite configuration of User1:
'    cfgHeadline = "R \r|'chapter£' R \r|R \r'£-£'R"     ' number | text | number - text
'    cfgBookmark = "R"                                   ' text   | pagenumber
'    cfgFigureTE = "R \r"                                ' Fig XX | description | combi

    ' Here you can define additional default parameters which shall generally be appended:
    ' Here we define
    '   - that cross references shall always be inserted as hyperlinks:
    '   - that the /* MERGEFORMAT switches shall be set
    cfgHeadline = AddDefaults(cfgHeadline, "\h \* MERGEFORMAT")
    cfgBookmark = AddDefaults(cfgBookmark, "\h \* MERGEFORMAT")
    cfgFigureTE = AddDefaults(cfgFigureTE, "\h \* MERGEFORMAT")
    '
    ' Define here the subtitles that shall be recognised. Add more as you wish:
    Const subtitleTypes = "Figure|Fig.|Abbildung|Abb.|Table|Tab.|Tabelle|Equation|Eq.|Gleichung"
    '
    ' Use regex-Syntax to define how to determine subtitles from headers:
    ' ("£" is a special character that will be replaced with the above <subtitleTypes>.)
    Const subtitleRecog = "((^(£))([\s\xa0]+)(\d+):?([\s\xa0]+)(.*))"
    ' Above example:
    '   To be recognised as a subtitle the string
    '      - must start with one of the keywords in <subtitlTypes>
    '      - be followed by one or more of (whitespaces or character xa0=160=&nbsp;)
    '      - be followed by one or more digits
    '      - be followed by zero or one colon
    '      - be followed by one or more of (whitespaces or character xa0=160=&nbsp;)
    '      - be followed by zero or more additional characters
    '
    ' === End of Configuration
    ' ============================================================================================
    
    
    ' === Initialisations ========================================================================
    cfgHeadline = Replace(cfgHeadline, "£", Chr$(160))
    cfgBookmark = Replace(cfgBookmark, "£", Chr$(160))
    cfgFigureTE = Replace(cfgFigureTE, "£", Chr$(160))
    cfgAHeadline = Split(CStr(cfgHeadline), "|")
    cfgABookmark = Split(CStr(cfgBookmark), "|")
    cfgAFigureTE = Split(CStr(cfgFigureTE), "|")
    linktypes = Split(subtitleTypes, "|")
    
    ActiveWindow.View.ShowFieldCodes = False
    
    'Debug.Print cfgPHeadline
    ' Stelle, wo die Referenz eingefügt werden soll:
    ' ============================================================================================
    ' === Check if we are in Insertion-Mode or not ===============================================
    If Not (isActive) Then
        ' ========================================================================================
        ' ===== We are NOT in Insertion-Mode!
        If ActiveDocument.Bookmarks.Exists("tempforInsert") Then
            ActiveDocument.Bookmarks.Item("tempforInsert").Delete
        End If
        
        ' Special function: if the cursor is inside a wdFieldRef-field, then
        ' - toggle the display among the configured options
        ' - remember the new status for future inserts.
        index = CursorInField(Selection.Range) ' would fail, if .View.ShowFieldCodes = True
        If index <> 0 Then
            ' ====================================================================================
            ' ===== Toggle display:
            Dim myOption As String
            Dim myRefType As Integer
            Dim fText0 As String                  '
            Dim fText2 As String                  ' Refnumber
            Dim element As Variant
            Dim needle As String
            Dim optionstring As String
            Dim idx As Integer
            Dim bmname As String
            
            ' == Read and clean the code from the field:
            fText0 = ActiveDocument.Fields(index).Code ' Original
            fText2 = fText0
            fText2 = Replace(fText2, "PAGE", "")        ' change from PAGEREF to REF
            fText2 = RegEx(fText2, "REF\s+(\S+)")       ' get the reference-name
            needle = Replace(subtitleRecog, "£", subtitleTypes)
            
            Select Case True
                Case Left(fText2, 4) = "_Ref" And isSubtitle(fText2, needle)
                    ' == It is a subtitle.
                    myRefType = wdRefTypeNumberedItem
                    'Debug.Print "Subtitle:", cfgPFigureTE, myOption
                    idx = MultifieldDelete(cfgAFigureTE, cfgPFigureTE, fText0, index, needle, True)
                    If idx = -1 Then Exit Sub
                        
                    cfgPFigureTE = (idx + 1) Mod (UBound(cfgAFigureTE) + 1)
                    myOption = cfgAFigureTE(cfgPFigureTE)
                    Application.StatusBar = "New Cross reference format for Subtitles: <" & myOption & ">."
                    
                    paramRefText = ActiveDocument.Bookmarks(fText2).Range.Paragraphs(1).Range.text
                    Call MultifieldDelete(cfgAFigureTE, cfgPFigureTE, fText0, index, needle, True)
                    paramRefType = RegExReplace(paramRefText, needle, "$2")
                    found = getXRefIndex(paramRefType, paramRefText, index)
                    Call InsertCrossRefs(1, myOption, paramRefType, index, , True)
                Case Left(fText2, 4) = "_Ref"
                    ' == It is a headline.
                    myRefType = wdRefTypeHeading
                    'Debug.Print "Headline:", cfgPHeadline, myOption
                    idx = MultifieldDelete(cfgAHeadline, cfgPHeadline, fText0, index)
                    If idx = -1 Then Exit Sub
                    
                    cfgPHeadline = (idx + 1) Mod (UBound(cfgAHeadline) + 1)
                    myOption = cfgAHeadline(cfgPHeadline)
                    Application.StatusBar = "New Cross reference format for Headlines: <" & myOption & ">."
                    Call InsertCrossRefs(2, myOption, myRefType, index, fText2, True)
                Case Else
                    ' == It is a bookmark
                    myRefType = wdRefTypeBookmark
                    'Debug.Print "Bookmark:", cfgBookmark, myOption
                    idx = MultifieldDelete(cfgABookmark, cfgPBookmark, fText0, index)
                    If idx = -1 Then Exit Sub
                    
                    cfgPBookmark = (idx + 1) Mod (UBound(cfgABookmark) + 1)
                    myOption = cfgABookmark(cfgPBookmark)
                    Application.StatusBar = "New Cross reference format for Bookmarks: <" & myOption & ">."

                    bmname = rgex(Trim(fText0), "(REF|PAGEREF)\s+(\S+)", "$2")
                    'fText2 = "dummy"
                    Call InsertCrossRefs(2, myOption, myRefType, fText2, fText2, True)
            End Select
            Exit Sub                              ' Finished changing the display of the reference.
            
        Else
            ' ====================================================================================
            ' ===== Insert temporary Bookmark:
            ' Remember the current position within the document by putting a bookmark there:
            ActiveDocument.Bookmarks.Add Name:="tempforInsert", Range:=Selection.Range
            isActive = True             ' remember that we are in Insertion-Mode
        End If
        
        ' Stelle, wo die zu referenzierende Stelle ist
    Else
        ' ================================
        ' ===== We ARE in Insertion-Mode!
        
        ' ===== Find out the type of the element to cross-reference to.
        '       It could be a Headline, Figure, Bookmark, ...
        paramRefType = ""
        Select Case Selection.Paragraphs(1).Range.ListFormat.ListType
            Case wdListSimpleNumbering              ' bullet lists, numbered Elements
                paramRefType = wdRefTypeNumberedItem
                paramRefKind = wdNumberRelativeContext
                paramRefText = Selection.Paragraphs(1).Range.ListFormat.ListString & _
                        " " & Trim(Selection.Paragraphs(1).Range.text)
                paramRefText = Replace(paramRefText, Chr(13), "")
                found = getXRefIndex(paramRefType, paramRefText, index)
                
            Case wdListOutlineNumbering             ' Headlines
                paramRefType = wdRefTypeHeading
                ' The following two lines of code fix a strange behaviour of Word (Bug?),
                '    see http://www.office-forums.com/threads/word-2007-insertcrossreference-wrong-number.1882212/#post-5869968
                paramRefType = wdRefTypeNumberedItem
                paramRefReal = "Headline"
                paramRefKind = wdNumberRelativeContext
                ' paramRefText = Selection.Paragraphs(1).Range.ListFormat.ListString
                ' Sometimes (probably in documents where things have been deleted with track changes), the command
                '    Selection.Paragraphs(1).Range.ListFormat.ListString
                ' doesn't work correctly. Therefore, we get the numbering differently:
                Dim oDoc As Document
                Dim oRange As Range
                Set oDoc = ActiveDocument
                Set oRange = oDoc.Range(Start:=Selection.Range.Start, End:=Selection.Range.End)
                'Debug.Print oRange.ListFormat.ListString
                paramRefText = oRange.ListFormat.ListString
                found = getXRefIndex(paramRefType, paramRefText, index)
                
            Case wdListNoNumbering                  ' SEQ-numbered items, Bookmarks and Figure/Table/Equation/...
                paramRefText = Trim(Selection.Paragraphs(1).Range.text)
                If (Selection.Paragraphs(1).Range.Fields.Count = 1) Then
                    If ((Selection.Paragraphs(1).Range.Fields(1).Type = wdFieldSequence) And _
                        (Left(Selection.Paragraphs(1).Range.Fields(1).Code.text, 8) = " SEQ Ref") And _
                        (Selection.Paragraphs(1).Range.Bookmarks.Count = 1)) Then
                        ' == a) SEQ-numbered item, a bibliographic reference à la <[32] Jackson, 1939, page 37>:
                        paramRefType = wdRefTypeBookmark
                        paramRefKind = wdContentText
                        paramRefReal = "Bookmark"
                        paramRefText = Selection.Paragraphs(1).Range.Bookmarks(1).Name
                        found = getXRefIndex(paramRefType, paramRefText, index)
                    Else
                        GoTo elseelse
                    End If
                Else
elseelse:
                    ' Bookmark or Figure/Table/Equation/...
                    paramRefText = Replace(paramRefText, Chr(30), "-") ' Bindestrich im Format "Figure 1-2" kommt komischerweise als chr(30) an, daher hier korrigieren
                    'paramRefText = Replace(paramRefText, Chr(160), " ")
                    paramRefText = Left(paramRefText, Len(paramRefText) - 1)  ' remove that strange hidden character at the end
                    For Each linktype In linktypes
                        If paramRefText Like linktype & "*" Then
                            ' == b) Figure/Table/...
                            paramRefReal = "FigureTE"
                            paramRefType = linktype
                            paramRefKind = wdOnlyLabelAndNumber
                            found = getXRefIndex(paramRefType, paramRefText, index)
                            Exit For
                        End If
                    Next
                    If paramRefType = "" Then
                        ' OK, it was not a Figure/Table/Equation/...
                        ' Let's check if we are in a bookmark:
                        
                        ' Bookmarks can overlap. Therefore we need an iteration.
                        ' For user experience, it is best if we select the innermost bookmark (= the shortest):
                        Dim bname As String
                        Dim bmlen As Variant
                        Dim bmlen2 As Long
                        bmlen = ""
                        For Each element In Selection.Bookmarks
                            bmlen2 = Len(element.Range.text)
                            If bmlen2 < bmlen Or bmlen = "" Then
                                bname = element.Name
                                bmlen = Len(element.Range.text)
                            End If
                        Next
                        If bmlen <> "" Then
                            ' == c) bookmark
                            paramRefReal = "Bookmark"
                            paramRefType = wdRefTypeBookmark
                            paramRefText = bname
                            found = getXRefIndex(paramRefType, paramRefText, index)
                        End If
                    End If
                End If
                
            Case Else                               ' Everything else
                ' nothing to do
        End Select                                  ' Now we know what element it is
    
        ' ===== Check, if we can cross-reference to this element:
        If paramRefType = "" Then
            ' Sorry, we cannot...
            prompt = "Hier ist nix Verlinkbares!" & Chr(10) & "Probier ' es sonstwo oder breche ab."
            prompt = "Cannot cross reference to this location." & vbNewLine & "Try elsewhere or abort."
            Response = MsgBox(prompt, 1)
            If Response = vbCancel Then
                Selection.GoTo What:=wdGoToBookmark, Name:="tempforInsert"
                If ActiveDocument.Bookmarks.Exists("tempforInsert") Then
                    ActiveDocument.Bookmarks.Item("tempforInsert").Delete
                End If
                isActive = False
            End If
            Exit Sub
        End If
        
    
        ' ===== Insert the cross-reference:
        retry = False
retryfinding:
        If (found = False) And (retry = False) Then
            ' Refresh, ohne dass es als Änderung getracked wird:
            storeTrackStatus = ActiveDocument.TrackRevisions
            ActiveDocument.TrackRevisions = False
            Selection.HomeKey Unit:=wdStory
            
            Do                                    ' alle SEQ-Felder abklappern
                lastpos = Selection.End
                Selection.GoTo What:=wdGoToField, Name:="SEQ"
                On Error Resume Next
                Debug.Print Err.Number
                allowed = False
                If paramRefType = wdRefTypeNumberedItem Then
                    allowed = True
                Else
                    If IsInArray(paramRefType, linktypes) Then
                        allowed = True
                        searchstring = " SEQ " & linktype
                        If Left(Selection.NextField.Code.text, Len(searchstring)) = searchstring Then
                            Selection.Fields.Update
                        End If
                    End If
                End If
                If allowed = False Then
                    MsgBox "sollte hier nie reinlaufen"
                End If
                
            Loop While (lastpos <> Selection.End)
            retry = True
            ActiveDocument.TrackRevisions = storeTrackStatus
            GoTo retryfinding
        End If
        
        ' Jetzt das eigentliche Einfügen des Querverweises an der ursprünglichen Stelle:
        Selection.GoTo What:=wdGoToBookmark, Name:="tempforInsert"
        If found = True Then
            ' Read the correct array the currently selected options:
            Select Case paramRefReal
                Case "Headline"
                    optionstring = cfgAHeadline(cfgPHeadline)
                    ' paramRefType = not 1, but 0
                Case "Bookmark"
                    optionstring = cfgABookmark(cfgPBookmark)
                    ' paramRefType = 2
                Case Else
                    optionstring = cfgAFigureTE(cfgPFigureTE)
                    ' paramRefType = 0
            End Select
            
            Call InsertCrossRefs(1, optionstring, paramRefType, index)
        Else
            myerrtxt = ""
            myerrtxt = vbNewLine & myerrtxt & "Interner Fehler"
            myerrtxt = vbNewLine & myerrtxt & "paramRefType = <" & paramRefType & ">"
            myerrtxt = vbNewLine & myerrtxt & "paramRefKind = <" & paramRefKind & ">"
            myerrtxt = vbNewLine & myerrtxt & "paramRefText = <" & paramRefText & ">"
            MsgBox "Interner Fehler"
            Stop
        End If
        isActive = False
        On Error Resume Next
        If ActiveDocument.Bookmarks.Exists("tempforInsert") Then
            ActiveDocument.Bookmarks.Item("tempforInsert").Delete
        End If
    End If
End Sub

Function AddDefaults(ByRef theString, tobeAdded As String) As String
    AddDefaults = RegExReplace(theString, "(R(EF)?|P(AGEREF)?)", "$1" & " " & tobeAdded)
End Function

Function InsertCrossRefs(mode As Integer, _
                         optionstring As String, _
                         ByVal paramRefType As Variant, _
                         index As Variant, _
                         Optional ByVal refcode As String = "", _
                         Optional moveCursor As Boolean = False)
    ' Parameters:
    '   <mode>=0    update by manipulating switches
    '         =1    insert via .InsertCrossReference
    '         =2    insert via .Fields.Add
    '   <optionstring>: current option (possibly with multiple parts and multiple switches)
    '   <paramRefType>: the Type of cross reference:
    '                       2 for bookmarks
    '                       0 for everything else
    '   <index>       : the index of source in Word's internal table or
    '                   the name of the bookmark
    '   <refcode>     : the reference's name, e.g. _REF6537428
    
    Dim i As Integer           '
    Dim thePart As Variant
    Dim isCode As Boolean
    Dim thePartOld As Variant
    Dim isCodeOld As Boolean
    Dim refcode2 As String
    Dim thePart2 As String
    Dim thePart3 As String
    
    thePartOld = ""
    i = 0
    If Len(optionstring) = 0 Then
        MsgBox "InsertCrossRefs detected a non-valid option: <" & optionstring & ">."
        Exit Function
    End If
    Do
        thePartOld = thePart
        isCodeOld = isCode
        i = i + 1
        ' Get the next part (there could be multiple...)
        thePart = GetPart(optionstring, i, isCode)
        If thePart = Error Then
            ' We have reached the last part!
            
            ' One thing before we return:
            ' If we got the parameter <moveCursor>
            ' (which will be the case if we have done a replacement, rather than a new
            ' insert) then we want the cursor positioned behind the last field, even if
            ' the last part was a text - like this the user can continue to toggle.
            ' Hence we have to move the cursor a bit back:
            If (moveCursor = True) And (Len(thePartOld) > 0) And (isCodeOld = False) Then
                Selection.MoveLeft wdCharacter, Len(thePartOld)
            End If
            Exit Do
        End If
        
        ' If it's a text, insert it
        If isCode = False Then
            Application.Selection.InsertAfter thePart
            Application.Selection.Move wdCharacter, 1
        Else
        ' It is a code sequence:
            ' When we modify with method = <0>, we have received a fieldcode <refcode>.
            ' We use this code to do the modification.
            ' If there are any additional insertions, these must be done with the .Fields.Add-method.
            ' Therefore, we have to prepare a proper fieldcode for that method.
            Call ReplaceAbbrev(thePart)
            If mode = 2 Then
                ' The complete code must be provided in refcode2. The other params are unused.
                refcode2 = " " & RegEx(thePart, "(PAGEREF|REF|P|R)")
                refcode2 = refcode2 & " " & refcode
                refcode2 = refcode2 & " " & rgex(CStr(thePart), "(PAGEREF|REF|P|R)(.*)", "$2")
            Else
                ' we must provide:
                ' 1) paramRefType
                '   nothing to do
                
                ' 2) index
                '   nothing to do
                
                ' 3) thePart3: Fieldcode w/o RefNr w/ switches
                '   nothing to do
                
                ' 4)
                refcode2 = "not used"
                
            End If
            
            If Insert1CrossRef(mode, paramRefType, index, thePart, refcode2) = False Then
                Exit Do
            End If
            
            If mode = 0 Then
                ' We have modified the first field.
                ' Any additional fields shall be inserted with the .Fields.Add-method.
                mode = 2
            End If
        End If

    Loop While True
    
End Function

Function Insert1CrossRef(mode As Integer, Optional param1 As Variant, _
                                          Optional param2 As Variant, _
                                          Optional param3 As Variant, _
                                          Optional param4 As Variant) As Boolean
    ' Parameter <mode>=0    update by manipulating switches     ==>
    '                 =1    insert via .InsertCrossReference
    '                       ==> param1: wdReferenceKind
    '                       ==> param2: wdReferenceItem / RefNr
    '                       ==> param3: Fieldcode w/o RefNr w/ switches
    '                       ==> param4: not used
    '                 =2    insert via .Fields.Add
    '                       ==> param1: not used
    '                       ==> param2: not used
    '                       ==> param3: not used
    '                       ==> param4: Fieldcode w/ RefNr w/ switches
    ' Returns: true : upon successful completion
    '          false: when an invalid paramter was detected
    Dim inclHyperlink As Boolean
    Dim inclPosition As Boolean
    Dim param0 As Variant
    Dim myCode As String
    Dim idx As Integer
    
    Select Case mode
        Case 0              ' update by manipulating switches
            With ActiveDocument.Fields(param2)
                If InStr(1, param3, "PAGEREF") Then
                    myCode = "PAGEREF "
                    param3 = Replace(param3, "PAGEREF", "")
                Else
                    myCode = "REF "
                End If
                myCode = myCode & param4 & " " & param3
                .Code.text = " " & myCode & " "
                 
                 ' If the cursor is now behind the field, it must be moved back:
                 If Selection.End > .Result.End Then
                     Selection.Move wdCharacter, -1
                 End If
                 Selection.Fields.Update
                 ' Now, the cursor will be exactly behind the field. That's fine.
                 
                 ' If the cursor is now in front of the field, it must be moved forward:
                 If Selection.Start < .Result.Start Then
                     Selection.Start = .Result.End
                 End If
            End With
        Case 1                  ' Insert new via .InsertCrossReference
            Dim mainSwitches() As String
            Dim mainFound As Boolean
            Dim element As Variant
            Dim rmatch As Variant
            ' Check the main switches:
            mainSwitches = Split("PAGEREF|P|REF|R", "|")
            mainFound = False
            For Each element In mainSwitches
                rmatch = RegEx(param3, "(\b" & Trim(element) & "\b)")
                If rmatch <> False Then
                'If InStr(1, param3, element) Then
                    mainFound = True
                    param3 = Trim(Replace(param3, rmatch, ""))
                    If Left(element, 1) = "R" Then
                        param3 = "REF " & param3
                        param0 = wdContentText
                        If Not (IsNumeric(param1)) Then param0 = wdOnlyCaptionText
                    Else
                        param3 = "PAGEREF " & param3
                        param0 = wdPageNumber
                    End If
                    Exit For
                End If
            Next
            If mainFound = False Then
                MsgBox "Insert1CrossRef: non-valid code encountered: <" & param3 & ">"
                Insert1CrossRef = False
                Exit Function
            End If
            
            ' Check the modifier switches:
            If InStr(1, param3, "\n") Then
                param0 = wdNumberNoContext
                param3 = Replace(param3, "\n", "")
            ElseIf InStr(1, param3, "\w") Then
                param0 = wdNumberFullContext
                param3 = Replace(param3, "\w", "")
            ElseIf InStr(1, param3, "\c") Then
                param0 = wdEntireCaption
                param3 = Replace(param3, "\c", "")
                If Not (IsNumeric(param1)) Then param0 = wdEntireCaption
            ElseIf InStr(1, param3, "\r") Then
                param0 = wdNumberRelativeContext
                param3 = Replace(param3, "\r", "")
                If Not (IsNumeric(param1)) Then param0 = wdOnlyLabelAndNumber
            End If
            If InStr(1, param3, "\h") Then
                inclHyperlink = True
                param3 = Replace(param3, "\h", "")
            Else
                inclHyperlink = False
            End If
            If InStr(1, param3, "\p") Then
                inclPosition = True
                param3 = Replace(param3, "\p", "")
            Else
                inclPosition = False
            End If
            '                                  RefType, RefKind, RefIndx, hyperlink,  position     sepNr , seperator
            Call Selection.InsertCrossReference(param1, param0, param2, inclHyperlink, inclPosition, False, "")
            param3 = Replace(param3, "PAGEREF", "")
            param3 = Replace(param3, "REF", "")
            
            ' Make sure, the cursor is still in the field
            Do
                idx = CursorInField(Selection.Range)
                If idx <> 0 Then Exit Do
                Selection.MoveLeft wdCharacter, 1
            Loop While True
            ' Append any leftover switches:
            ActiveDocument.Fields(idx).Code.text = ActiveDocument.Fields(idx).Code.text & " " & param3 & " "
            'Application.StatusBar = "Cross Reference inserted of type <" & param3 & ">."
        Case 2              ' Insert new via .Fields.Add
            Dim index As Integer
            
            Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False
            Selection.TypeText text:=Trim(param4)
            Selection.Fields.Update

            ' Put Cursor behind the new field:
            Selection.Move wdCharacter, 1
            Selection.Fields.Update
            'Application.StatusBar = "Cross Reference inserted <" & param4 & ">."
        Case Else
            Stop
    End Select
    
    Insert1CrossRef = True
End Function

Public Function IsInArray(ByVal stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
    
End Function

Function CursorInField(theRange As Range) As Long
    ' If the cursor is currently positioned in a Word field of type wdFieldRef,
    ' then this function returns the index of this field.
    ' Else it returns 0.

    Dim Item   As Variant

    CursorInField = 0
    'Debug.Print Selection.Start
    
    ' There is Selection.Fields or Range.Fields, which looks promising to find
    ' the field over which the cursor stands.
    ' But the fields are only listed if the range or selection overlaps them fully,
    ' not on partly overlap.
    ' Therefore, we just iterate over all the fields and check their start- and end-position
    ' against the position of the cursor.
    For Each Item In ActiveDocument.Fields
        'If Item.index < 50 Then Debug.Print Item.index, Item.Type, Item.Result.Start, Item.Result.End, Item.Result.Case
        If Item.Type = wdFieldRef Or Item.Type = wdFieldPageRef Then ' wdFieldRef:=3
            If Item.Result.Start <= theRange.Start And _
                 Item.Result.End >= theRange.Start - 1 Then ' -1 allows that the cursor may stand immediately behind the field
                CursorInField = Item.index
                'Debug.Print "CursorInField: yes"
                Exit Function
            End If
        End If
    Next
End Function

Private Sub test_CursorInField()
    Debug.Print IsCursorInField
End Sub

Function RegExReplace(Quelle As Variant, Expression As Variant, replacement As Variant) As String
    ' Beispiel für einen Aufruf:
    ' (würde bei mehrfachen Backslashes hintereinander jeweils den ersten wegnehmen)
    ' result = RegExReplace(input, "\\(\\+)", "$1")

    Dim re     As New RegExp

    re.Global = True
    re.Pattern = Expression
    RegExReplace = re.Replace(Quelle, replacement)
    
End Function

Function RegEx(Quelle As Variant, Expression As String) As Variant
    Dim re     As New RegExp
    Dim extract As Object
    Dim extract2 As Object
    Dim i As Integer

    re.Global = True
    re.Pattern = Expression
    Set extract = re.Execute(Quelle)
    On Error Resume Next
    RegEx = extract.Item(0).SubMatches.Item(0)
    If Error <> "" Then
        RegEx = False
    End If
    
End Function

Function rgex(strInput As String, matchPattern As String, _
  Optional ByVal outputPattern As String = "$0", Optional ByVal behaviour As String = "") As Variant
  ' How it works:
  ' It takes 2-3 parameters.
  '    A text to use the regular expression on.
  '    A regular expression.
  '    A format string specifying how the result should look. It can contain $0, $1, $2, and so on.
  '         $0 is the entire match, $1 and up correspond to the respective match groups in the regular expression.
  '         Defaults to $0.
  '    If the expression matches multiple times, by default only the first match ("0") is considered.
  '         This can be modified by the optional parameter.
  '         It can contain 1, 2, 3, ... for the 1st, 2nd, 3rd, ... match or "*" to return the complete array of matches.
  '
  ' Some examples
  ' Extracting an email address:
  ' =rgex("Peter Gordon: some@email.com, 47", "\w+@\w+\.\w+")
  ' =rgex("Peter Gordon: some@email.com, 47", "\w+@\w+\.\w+", "$0")
  ' Results in: some@email.com
  ' Extracting several substrings:
  ' =rgex("Peter Gordon: some@email.com, 47", "^(.+): (.+), (\d+)$", "E-Mail: $2, Name: $1")
  ' Results in: E-Mail: some@email.com, Name: Peter Gordon
  ' To take apart a combined string in a single cell into its components in multiple cells:
  ' =rgex("Peter Gordon: some@email.com, 47", "^(.+): (.+), (\d+)$", "$" & 1)
  ' =rgex("Peter Gordon: some@email.com, 47", "^(.+): (.+), (\d+)$", "$" & 2)
  ' Results in: Peter Gordon some@email.com ...
  '
  ' Prerequisites: Verweis auf
  ' Microsoft VBScript Regular Expressions 5.5  |c:\windows\SysWOW64\vbscript.dll\3
  '
  ' Modified from source: https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops/22542835

  Dim inputRegexObj As New VBScript_RegExp_55.RegExp, outputRegexObj As New VBScript_RegExp_55.RegExp, outReplaceRegexObj As New VBScript_RegExp_55.RegExp
  Dim inputMatches As Object, replaceMatches As Object, replaceMatch As Object
  Dim replaceNumber As Integer
  Dim ixi As Integer
  Dim ixf As Integer
  Dim ix As Integer
  Dim sepr As String
  Dim outputres As String
  
  With inputRegexObj
      .Global = True
      .MultiLine = True
      .IgnoreCase = False
      .Pattern = matchPattern
  End With
  With outputRegexObj
      .Global = True
      .MultiLine = True
      .IgnoreCase = False
      .Pattern = "\$(\d+)"
  End With
  With outReplaceRegexObj
      .Global = True
      .MultiLine = True
      .IgnoreCase = False
  End With

  Select Case behaviour
    Case "":                          ' No parameter given: by default use match 0
      ixi = 0   ' index initial
      ixf = ixi ' index final
    Case IsNumeric(behaviour):        ' They want a specific match
      ixi = CInt(behaviour) - 1       ' => 1st match has index 0
      ixf = ixi
    Case "*":                         ' They want all matches
      ixi = 0
      ixf = -1  ' preliminary value
    Case Else
      MsgBox "errorhasoccured"
  End Select

  Set inputMatches = inputRegexObj.Execute(strInput)
  If inputMatches.Count = 0 Then                  ' Nothing found
    rgex = False
  ElseIf (ixi + 1 > inputMatches.Count) Then      ' There is no x-th match-group
    rgex = False
  Else                                            ' Something was found
    rgex = ""
    sepr = ""
    If ixf = -1 Then                              ' Now we can determine, how many matches to return
      ixf = inputMatches.Count - 1
      ' Outputformat will be: "{Nr of results}|{Result 1}|{Result 2}|..|{Result N}"
      sepr = "|"
      rgex = CStr(inputMatches.Count)
    End If
    
    ' Reduce results to the requested match-group:
    Set replaceMatches = outputRegexObj.Execute(outputPattern)
    
    For ix = ixi To ixf
      For Each replaceMatch In replaceMatches
        replaceNumber = replaceMatch.SubMatches(0)
        outReplaceRegexObj.Pattern = "\$" & replaceNumber
    
        If replaceNumber = 0 Then
          outputres = outReplaceRegexObj.Replace(outputPattern, inputMatches(ix).Value)
        Else
          If replaceNumber > inputMatches(ix).SubMatches.Count Then
            'rgex = "A to high $ tag found. Largest allowed is $" & inputMatches(0).SubMatches.Count & "."
            rgex = Error 'CVErr(vbErrValue)
            Exit Function
          Else
            outputres = outReplaceRegexObj.Replace(outputPattern, inputMatches(ix).SubMatches(replaceNumber - 1))
          End If
        End If
      Next
      rgex = rgex & sepr & outputres
    Next
  End If
End Function

Function isSubtitle(bookmark As String, regexneedle As String) As Boolean
    Dim theText As String

    isSubtitle = False
    
    If ActiveDocument.Bookmarks.Exists(bookmark) = False Then
        Exit Function
    End If
    
    theText = ActiveDocument.Bookmarks(bookmark).Range.Paragraphs(1).Range.text
    theText = Replace(theText, Chr(160), " ")
    If RegEx(theText, regexneedle) <> False Then
        isSubtitle = True
    End If
    
End Function

Function MultifieldDelete(ByRef optionArray() As String, _
                          ByRef optionPtr As Integer, _
                          myCode As String, _
                          ByRef index, _
                          Optional checkbyText As String = "", _
                          Optional includeLast As Boolean = False) As Integer
    ' Returns: -1 on error
    '          else the index of the found format
    ' The Parameters:
    '   optionArray  (in): Array of Options
    '   optioPtrr    (IO): ptr to the current Option within the array
    '   myCode           : not used
    '   index        (in): index of the field or name of the bookmark
    '   checkbyText  (in): for the type FigureTE
    '   includeLast  (in): not used
    
    Dim i, j     As Integer         ' loop over options and parts
    Dim theCode As String           ' the field's code
    Dim myRange As Range            ' a moving range used to verify the content of each part
    Dim theOption As String         ' the current option
    Dim thePart As String           ' the current part
    Dim isCode As Boolean           ' whether the part is field code or text
    Dim matchfound As Boolean       ' to keep track of the check results
    Dim idx As Integer              ' index of the current field in Word's array
    Dim theEnd As Long              ' to remember the end of the current multithing
    Dim lOptionPtr As Integer       ' a local copy of the optionPtr, which we increment while searching
    Dim textOK As Boolean
    Dim fulltext As String
    Dim expctd As String
    Dim theText As String
    
    MultifieldDelete = -1
    
    ' Create a dummy Range:
    Set myRange = ActiveDocument.Range
    
    ' Loop over the options:
    For i = 0 To UBound(optionArray)
        lOptionPtr = (optionPtr + i) Mod (UBound(optionArray) + 1)
        theOption = optionArray(lOptionPtr)
        'Debug.Print theOption
        If Len(theOption) = 0 Then
            MsgBox "MultifieldDelete has encountered an invalid option <" & theOption & ">."
            Exit Function
        End If
        
        ' Loop over the parts:
        j = 0
        Do While True
            matchfound = True
            j = j - 1
            thePart = GetPart(theOption, j, isCode)
            If j = 0 Then
                ' There is no next part, so exit
                Exit Do
            End If
            If thePart = Error Then
                ' We have checked all parts
                Exit Do
            End If
            
            If j = -1 Then
                ' Make sure, the Cursor is immediately behind the field:
                ActiveDocument.Fields(index).Update
                Set myRange = Selection.Range
                
                ' If it is a multipart thingy, include the last text in our Range:
                If isCode = False Then
                    myRange.MoveEnd wdCharacter, Len(thePart)
                    myRange.Start = myRange.End
                End If
                
                ' Remember the end of the multithing:
                theEnd = myRange.End
                
            End If
            
            If isCode = False Then
                myRange.MoveStart wdCharacter, -Len(thePart)
                If myRange.text = thePart Then
                    ' OK, matched
                Else
                    ' Mismatch
                    matchfound = False
                    Exit Do
                End If
                myRange.MoveEnd wdCharacter, -Len(thePart)
            Else
                ' It is a Code
                ' Check, if there is a field:
                Call ReplaceAbbrev(thePart)
                idx = CursorInField(myRange)
                If idx = 0 Then
                    matchfound = False
                    Exit Do
                End If
                
                If checkbyText <> "" Then
                    ' This is for the type FigureTE.
                    ' Here, the switches \r, \c are not applicable,
                    ' rather the Ref points to different bookmarks.
                    fulltext = ActiveDocument.Bookmarks(RegEx(myCode, "(_Ref\d+)")).Range.Paragraphs(1).Range.text
                    textOK = False
                    If InStr(1, thePart, "PAGEREF") Then
                        textOK = True   ' because there is no real check
                    ElseIf InStr(1, thePart, "\p") Then         ' above/below
                        textOK = True   ' because there is no real check
                    ElseIf InStr(1, thePart, "\r") Then
                        ' Category and number
                        thePart = Trim(Replace(thePart, "\r", ""))
                        expctd = RegExReplace(fulltext, checkbyText, "$3$4$5")
                    ElseIf InStr(1, thePart, "\c") Then
                        ' Full subtitle
                        thePart = Trim(Replace(thePart, "\c", ""))
                        expctd = fulltext
                    Else
                        ' Description only
                        expctd = RegExReplace(fulltext, checkbyText, "$7")
                    End If
                    If textOK = False Then
                        theText = ActiveDocument.Fields(idx).Result.text
                        expctd = Replace(expctd, Chr$(13), "")
                        If StrComp(theText, expctd, vbTextCompare) = 0 Then
                            textOK = True
                        End If
                    End If
                    'Else
                        theCode = ActiveDocument.Fields(idx).Code.text
                        theCode = Trim(RegExReplace(theCode, "(REF|PAGEREF)\s+(\S+)", "$1")) ' Remove the bookmark name
                    'End If
                Else
                    textOK = True       ' because there is no check
                    theCode = ActiveDocument.Fields(idx).Code
                    theCode = Trim(RegExReplace(theCode, "(REF|PAGEREF)\s+(\S+)", "$1")) ' Remove the bookmark name
                End If
                
                ' Check, if the Field codes match (present code vs expected code):
                If CodesComply(theCode, thePart) = False Or textOK = False Then
                    matchfound = False
                    Exit Do
                End If
                myRange.MoveStart wdCharacter, -Len(ActiveDocument.Fields(idx).Result.text)
                myRange.MoveEnd wdCharacter, -Len(ActiveDocument.Fields(idx).Result.text)
            End If
            
        Loop    ' over the parts
        If matchfound Then
            MultifieldDelete = lOptionPtr
            optionPtr = lOptionPtr
            Exit For     ' no need to check the other options
        End If
    Next        ' over the options
    
    If matchfound = False Then
        ' Not successful in finding the pattern
        Exit Function
    End If
    
    If Abs(j) > 1 Then
        ' It was a multifield
    Else
        ' It was a single field
    End If
    
    ' Delete the whole thing:
    myRange.End = theEnd
    Dim theStart As Long
    theStart = myRange.Start
    myRange.Cut
    ' Because Word may try to be smart by removing a lonely blank:
    If Selection.Start < theStart Then
        Selection.InsertBefore (" ")
        Selection.Move wdCharacter, 1
    End If
    
End Function

Function GetPart(theString, thePosition, Optional ByRef isCode As Boolean) As Variant
    ' thePosition counts  1, 2, ...
    '                 or -1, -2, ... to find from behind
    ' It is an I/O-parameter and will be reset to 0 in case of error.

    Dim re As New RegExp
    Dim extract As Object
    Dim idx As Integer
    
    ' 1) Get the different parts
    theString = Trim(theString)
    re.Global = True
    re.Pattern = "[^']+(?='|$)" '"[^'|$]+(?='|^)"
    Set extract = re.Execute(theString)
    
    ' 2) Check if the index is out of bounds:
    If Abs(thePosition) > extract.Count Then
        thePosition = 0
        GetPart = ""
        Exit Function
    End If
    If thePosition = 0 Then
        MsgBox "GetPart() has received invalid index <" & thePosition & ">."
        Stop
    End If
    If thePosition < 0 Then
        ' Find from behind
        idx = extract.Count + thePosition
    Else
        idx = thePosition - 1
    End If
    
    ' 3) Extract the desired item
    GetPart = extract.Item(idx)
    If Len(GetPart) = 0 Then
        MsgBox "GetPart() has encountered invalid part: <" & GetPart & ">."
        Stop
    End If
    
    ' 4) As an additional information, return whether this is a string or a code
    If (Left(theString, 1) = "'") Then
        isCode = False
    Else
        isCode = True
    End If
    If ((idx Mod 2) > 0) Then
        isCode = Not (isCode)
    End If
        
End Function

Function ReplaceAbbrev(theString) As Boolean
    Dim mainFound As Boolean
    Dim element As Variant
    
    Dim rmatch As Variant
    Dim needle As String
    Dim repl As String
    
    ReplaceAbbrev = False
    
    needle = "(PAGEREF|P|REF|R)\b"          '\b is for word boundary
    'needle = "(OKKLJLK|O|ZUI|Z)\b"
    rmatch = RegEx(theString, needle)
    If rmatch = False Then
        MsgBox "Expected keyword not found in <" & theString & ">."
        Exit Function
    End If
    
    If Left(rmatch, 1) = "P" Then
        repl = "PAGEREF"
    Else
        repl = "REF"
    End If
    theString = RegExReplace(theString, needle, repl)
    theString = Trim(theString)
    
    ReplaceAbbrev = True
    
End Function

Function CodesComply(ByVal CodeToBCheck As String, ByVal CodeExpected As String) As Boolean
    Dim thisSwitch As String
    Dim element As Variant

    CodeToBCheck = Trim(CodeToBCheck)
    CodeExpected = Trim(CodeExpected)
    Call ReplaceAbbrev(CodeExpected)
    
    ' Code complies, if there are exactly the same elements. Order is arbitrary.
    ' Extract the individual words with a regex:
    Dim re As New RegExp
    Dim extract As Object
    re.Global = True
    re.Pattern = "\S+"                  ' Word by Word
    Set extract = re.Execute(CodeExpected)
    
    For Each element In extract
        If InStr(1, CodeToBCheck, element) > 0 Then
            CodeToBCheck = Trim(Replace(CodeToBCheck, element, ""))
        Else
            CodesComply = False
            Exit Function
        End If
    Next
    
    ' If there is nothing leftover now in theCode except "REF", we have a match:
    If Len(CodeToBCheck) > 0 Then
        CodesComply = False
        Exit Function
    Else
        CodesComply = True
    End If
    
End Function

Function getXRefIndex(RefType, text, index As Variant) As Boolean
        
    Dim thisitem As String
    Dim element As Variant
    Dim i As Integer
    
    text = Trim(text)
    If Right(text, 1) = Chr$(13) Then text = Left(text, Len(text) - 1)
    
    getXRefIndex = False
    If RefType = wdRefTypeBookmark Then
        ' The "index" is the bookmark name:
        index = text
        getXRefIndex = True
    Else
        ' In all other cases, we need to find the index
        ' by searching through Word's CrossReferenceItems(RefType):
        index = -1
        For i = 1 To UBound(ActiveDocument.GetCrossReferenceItems(RefType))
            thisitem = Trim(Left(Trim(ActiveDocument.GetCrossReferenceItems(RefType)(i)), Len(text)))
            text = Replace(text, Chr(160), " ")
            If StrComp(thisitem, text, vbTextCompare) = 0 Then
                getXRefIndex = True
                index = i
                Exit For
            End If
        Next
    End If

End Function

