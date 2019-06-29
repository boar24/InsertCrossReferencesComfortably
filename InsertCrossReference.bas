Attribute VB_Name = "InsertCrossReference"
Option Explicit

Sub CreateCrossReference()
'
' Macro to insert cross references comfortably
'
' Preparation:
' 1) Put this code in a VBA module in your document or document template. 
'    It is recommended to put it into normal.dot, then the function is available in every document.
' 2) Assign a keyboard shortcut to this macro (recommendation: AltGr-Q)
'    This works like
'     File -> Options -> Adapt Ribbon -> Keyboard Shortcuts: Modify...
'     Categories: Macros -> Macros: [select name of Macro] -> ...
'
' Useage:
' 1) At the location in the document, where the crossreference shall be inserted,
'    press the keyboard shortcut.
'    A temporary bookmark is inserted (if their display is enabled, grey square brackets will appear).
' 2) Move the cursor to the location to where the crossref shall point.
'    Supported are:
'    * Headlines
'    * Subtitles of Figures realised via { SEQ Figure}, e.g. "Figure 123", "Figure 12-345"
'    * Subtitles of Tables  realised via { SEQ Table} , e.g. "Table 123", "Table 12-345"
'    * References to documents realised via { SEQ Ref}, e.g. "[42]"
'    Recommendation for large documents: use the navigation pane (View -> Navigation -> Headlines) 
'    Hint: Cross references to hidden text are not possible
'    Hint: The macro may fail trying to cross reference to locations that have heavily been edited 
'          (deletions / moves) with "track changes" (markup mode) turned on. 
' 3) Press the keyboard shortcut again.
'    The cursor will jump back to the location of insertion 
'    and the crossref will be inserted. Done!
' 4) Additional function:
'    By default, numerical references are inserted (e.g. "Figure 123"). 
'    When you press the keyboard shortcut when the cursor is already in a cross reference field,
'    - that field is toggled between <numerical reference> and <text reference> (e.g. "Overview")
'    - subsequently added cross references will use the latest format (persistent until closure of Word)
'
' Revision History:
' 151204 Beginn der Revision History
' 160111 Kann jetzt auch umgehen mit Numerierungen mit Bindestrich à la "Figure 1-1"
' 160112 Jetzt auch Querverweise möglich auf Dokumentenreferenzen à la "[66]" mit Feld " SEQ Ref "
' 160615 Felder werden upgedatet falls nötig
' 180710 Support für "Nummeriertes Element"
' 181026 Generischerer Code für Figure¦Table¦Abbildung
' 190628 New function: toggle to insert numeric or text references ("\r")
' 190629 Explanations and UI changed to English

  Static isActive As Boolean
  Static isTextRef As Boolean
  
  Dim type1, type2, type3, type4, type5 As Variant
  Dim Response, storeTrackStatus, lastpos As Variant
  Dim prompt As String
  Dim retry, found As Boolean
  Dim index, slaNumItems, ilRefItem  As Integer
  Dim thisitem, myerrtxt As String
  Dim linktype, searchstring As String
  Dim linktypes() As String
  Dim allowed As Boolean
  
  linktypes = Split("Figure,Table,Abbildung", ",")
  
  ' Stelle, wo die Referenz eingefügt werden soll:
  If Not (isActive) Then
    On Error Resume Next
    ActiveDocument.Bookmarks.Item("tempforInsert").Delete
    
    ' Special function: if the cursor is inside a wdFieldRef-field, then
    ' - toggle the parameter <\r > (remove/insert). If removed, the name is displayed rather than the number.
    ' - remember the new status for future inserts.
    index = CursorInField()
    If index <> 0 Then
        With ActiveDocument.Fields(index).Code
            If InStr(1, .Text, "\r") Then
                ' entfernen
                .Text = Replace(.Text, "\r \h", "\h")
                isTextRef = True
            Else
                ' einfügen
                .Text = Replace(.Text, "\h", "\r \h")
                isTextRef = False
            End If
            Selection.Fields.Update
        End With
        Exit Sub
        
    Else
        ' Remember the current position within the document by putting a bookmark
        ActiveDocument.Bookmarks.Add Name:="tempforInsert", Range:=Selection.Range
        ' Go into Insertion-mode:
        isActive = True
    End If
    
  ' Stelle, wo die zu referenzierende Stelle ist
  Else
    Call refreshIt
    ' Typ herausfinden (Überschriften, Figures, Dokumentenreferenzen, ...):
    type2 = ""
    Select Case Selection.Paragraphs(1).Range.ListFormat.ListType
      Case wdListOutlineNumbering
        ' Headlines / Überschriften
        type1 = wdRefTypeHeading
        type2 = "Überschrift"
' folgende 2 Zeilen fixen ein komisches Verhalten (Bug?) in Word, siehe http://www.office-forums.com/threads/word-2007-insertcrossreference-wrong-number.1882212/#post-5869968
        type1 = wdRefTypeNumberedItem
        type2 = "Nummeriertes Element"
        type3 = wdNumberRelativeContext
        ' type4 = Selection.Paragraphs(1).Range.ListFormat.ListString
        ' Sometimes (probably in documents where things have been deleted with track changes), the command
        '    Selection.Paragraphs(1).Range.ListFormat.ListString
        ' doesn't work correctly. Therefore, we get the numbering differently:
        Dim oDoc As Document
        Dim oRange As Range
        Set oDoc = ActiveDocument
        Set oRange = oDoc.Range(Start:=Selection.Range.Start, End:=Selection.Range.End)
        'Debug.Print oRange.ListFormat.ListString
        type4 = oRange.ListFormat.ListString
    
        type5 = Len(type4)
        
      Case wdListNoNumbering
        ' darunter fallen: Referenzen, "Figure", "Table":
        type4 = Trim(Selection.Paragraphs(1).Range.Text)
        If ((Selection.Paragraphs(1).Range.Fields.Count = 1) And _
            (Selection.Paragraphs(1).Range.Fields(1).Type = wdFieldSequence) And _
            (Left(Selection.Paragraphs(1).Range.Fields(1).Code.Text, 8) = " SEQ Ref") And _
            (Selection.Paragraphs(1).Range.Bookmarks.Count = 1) _
           ) Then
          ' Dokumentenreferenzen
          type1 = wdRefTypeBookmark
          type2 = "Textmarke"
          type3 = wdContentText
          type4 = Selection.Paragraphs(1).Range.Bookmarks(1).Name
          type5 = Len(type4)
        Else
          ' z.B. Figures, aber auch Fliesstext
          type4 = Replace(type4, Chr(30), "-")      ' Bindestrich im Format "Figure 1-2" kommt komischerweise als chr(30) an, daher hier korrigieren
          type5 = Len(type4) - 1                    ' irgendwie klebt da noch ein komischer hidden character am schluss...
          type4 = Left(type4, type5)
          For Each linktype In linktypes
            If Left(type4, Len(linktype) + 1) = linktype & " " Then
              type1 = linktype
              type2 = type1
              type3 = wdOnlyLabelAndNumber
              Exit For
            End If
          Next
        End If
        
      Case wdListSimpleNumbering
      ' darunter fallen: Bulletlists, numberierte Elemente
        type1 = wdRefTypeNumberedItem
        type2 = "Nummeriertes Element"
        type3 = wdNumberRelativeContext
        type4 = Selection.Paragraphs(1).Range.ListFormat.ListString & _
                " " & Trim(Selection.Paragraphs(1).Range.Text)
        type4 = Replace(type4, Chr(13), "")
        type5 = Len(type4)
        
        ' Bulletlists
      Case Else
        ' Sonstwas
    End Select
    
    If type2 = "" Then
      prompt = "Hier ist nix Verlinkbares!" & Chr(10) & "Probier' es sonstwo oder breche ab."
			prompt = "Cannot crossreference to this location!" & Chr(10) & "Try elsewhere or abort."
      Response = MsgBox(prompt, 1)
      If Response = vbCancel Then
        Selection.GoTo What:=wdGoToBookmark, Name:="tempforInsert"
        On Error Resume Next
        ActiveDocument.Bookmarks.Item("tempforInsert").Delete
        isActive = False
      End If
    Else
      retry = False
retryfinding:
      If type1 = wdRefTypeBookmark Then
        index = type4
        found = True
      Else
        ' In den anderen Fällen kann nicht direkt mit dem Namen reingegangen werden, sondern wir müssen den Index ermitteln:
        slaNumItems = ActiveDocument.GetCrossReferenceItems(type1)
        found = False
        index = 0
        For ilRefItem = 1 To UBound(slaNumItems)
          thisitem = Trim(slaNumItems(ilRefItem))
          thisitem = Trim(Left(thisitem, type5))
          
          'Debug.Print Len(thisitem), Len(type4)
          'For i = 1 To 32
          '  Debug.Print i, Mid(thisitem, i, 1), Asc(Mid(thisitem, i, 1)), Mid(type4, i, 1), Asc(Mid(type4, i, 1))
          'Next
'If ilRefItem = 84 Then Stop
          If StrComp(thisitem, Trim(type4), vbTextCompare) = 0 Then
            found = True
            index = ilRefItem
            Exit For
          End If
        Next
      End If
      If (found = False) And (retry = False) Then
        ' Refresh, ohne dass es als Änderung getracked wird:
        storeTrackStatus = ActiveDocument.TrackRevisions
        ActiveDocument.TrackRevisions = False
        Selection.HomeKey Unit:=wdStory
        
        Do        ' alle SEQ-Felder abklappern
          lastpos = Selection.End
          Selection.GoTo What:=wdGoToField, Name:="SEQ"
          On Error Resume Next
          Debug.Print Err.Number
          allowed = False
          If type1 = wdRefTypeNumberedItem Then
            'Selection.GoTo What:=wdGotoFigure, Count:=ilRefItem
            allowed = True
          Else
            If IsInArray(type1, linktypes) Then
              allowed = True
              searchstring = " SEQ " & linktype
              If Left(Selection.NextField.Code.Text, Len(searchstring)) = searchstring Then
                Selection.Fields.Update
              End If
            End If
          End If
          If allowed = False Then
            MsgBox "Code should never get here. Else, inform the developer."
          End If

        Loop While (lastpos <> Selection.End)
        retry = True
        ActiveDocument.TrackRevisions = storeTrackStatus
        GoTo retryfinding
      End If
      
      ' Jetzt das eigentliche Einfügen des Querverweises an der ursprünglichen Stelle:
      Selection.GoTo What:=wdGoToBookmark, Name:="tempforInsert"
      If found = True Then
      
        ' Gegebenenfalls überschreiben wir type3:
        If isTextRef Then
            type3 = wdContentText
        End If
        
        Selection.InsertCrossReference _
           ReferenceType:=type2, _
           ReferenceKind:=type3, _
           ReferenceItem:=index, _
           InsertAsHyperlink:=True, _
           IncludePosition:=False, _
           SeparateNumbers:=False, _
           SeparatorString:=" "
        'Selection.Fields.Update
      Else
        myerrtxt = ""
        myerrtxt = vbNewLine & myerrtxt & "Internal Error"
        myerrtxt = vbNewLine & myerrtxt & "type1 = <" & type1 & ">"
        myerrtxt = vbNewLine & myerrtxt & "type2 = <" & type2 & ">"
        myerrtxt = vbNewLine & myerrtxt & "type3 = <" & type3 & ">"
        myerrtxt = vbNewLine & myerrtxt & "type4 = <" & type4 & ">"
        myerrtxt = vbNewLine & myerrtxt & "type5 = <" & type5 & ">"
        MsgBox "Internal Error. Please insert cross reference manually."
        Stop
      End If
     isActive = False
     On Error Resume Next
     ActiveDocument.Bookmarks.Item("tempforInsert").Delete
    End If
  End If
End Sub

Private Sub refreshIt()
  Selection.HomeKey Unit:=wdLine
  Selection.EndKey Unit:=wdLine, Extend:=wdExtend
  Selection.Fields.Update
  Selection.HomeKey Unit:=wdLine
End Sub

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

Function CursorInField() As Long
    ' If the cursor is currently positioned in a Word field of type wdFieldRef,
    ' then this function returns the index of this field.
    ' Else it returns 0.
    
    Dim Item As Variant
    
    CursorInField = 0
    'Debug.Print Selection.Start
    For Each Item In ActiveDocument.Fields
        If Item.Type = wdFieldRef Then  ' wdFieldRef:=3
            If Item.Result.Start < Selection.Start And _
               Item.Result.End > Selection.Start Then
               CursorInField = Item.index
               Exit Function
            End If
        End If
    Next
End Function

Sub test_CursorInField()
    Debug.Print IsCursorInField
End Sub
