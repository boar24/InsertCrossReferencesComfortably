# InsertCrossReferencesComfortably
**VBA macro to comfortably insert cross references in MS Word**  

When inserting cross references in MS Word, have you ever been annoyed how tedious it is to do via the miniature default dialog (see screenshot below)? One has to  
- Navigate through the ribbon to find Insert -> Links -> Cross Reference (only visible if the Word window is very wide) 
- Click it
- Select what shall be cross-referenced
- Select how the cross-reference shall be displayed
- Find the element in the list (which, prior to Office 2016 is diminutive and fixed in height !) and click it

How about that instead:
- Press a hotkey when you want to insert a crossreference at the current location
- Put the cursor to the element to which the crossreference shall point to, e.g. a figure label or a headline
- Press the hotkey again => the cursor jumps back to the original location and the crossref is inserted

  
![kdkdk](https://github.com/Traveler4/InsertCrossReferencesComfortably/blob/master/Zwischenablage01.png)

  
## Preparation:  
1) Put this code in a VBA module in your document or document template.   
   It is recommended to put it into normal.dot, then the function is available in every document.  
2) Assign a keyboard shortcut to this macro (recommendation: AltGr-Q)  
   This works like  
    File -> Options -> Adapt Ribbon -> Keyboard Shortcuts: Modify...   
    Categories: Macros -> Macros: [select name of Macro] -> ...  
## Useage:  
1) At the location in the document, where the crossreference shall be inserted,  
   press the keyboard shortcut.  
   A temporary bookmark is inserted (if their display is enabled, grey square brackets will appear).  
2) Move the cursor to the location to where the crossref shall point.  
   Supported are:  
   * Headlines  
   * Subtitles of Figures realised via { SEQ Figure}, e.g. "Figure 123", "Figure 12-345"  
   * Subtitles of Tables  realised via { SEQ Table} , e.g. "Table 123", "Table 12-345"  
   * References to documents realised via { SEQ Ref}, e.g. "[42]"  
   _Recommendation for large documents:_ use the navigation pane (View -> Navigation -> Headlines)   
3) Press the keyboard shortcut again.  
   The cursor will jump back to the location of insertion   
   and the crossref will be inserted. Done!  
4) Additional function:  
   By default, numerical references are inserted (e.g. "Figure 123").   
   When you press the keyboard shortcut when the cursor is already in a cross reference field,  
   - that field is toggled between <numerical reference> and <text reference> (e.g. "Overview")  
   - subsequently added cross references will use the latest format (persistent until Word is exited)  
  
Limitations:
  * Cross references to hidden text are not possible  
  * The macro may fail trying to cross reference to locations that have heavily been edited (deletions / moves) with "track changes" (markup mode) turned on.   

## Revision History:  
* 151204 Beginn der Revision History  
* 160111 Kann jetzt auch umgehen mit Numerierungen mit Bindestrich à la "Figure 1-1"  
* 160112 Jetzt auch Querverweise möglich auf Dokumentenreferenzen à la "[66]" mit Feld " SEQ Ref "  
* 160615 Felder werden upgedatet falls nötig  
* 180710 Support für "Nummeriertes Element"  
* 181026 Generischerer Code für Figure¦Table¦Abbildung  
* 190628 New function: toggle to insert numeric or text references ("\r")  
* 190629 Explanations and UI changed to English  
