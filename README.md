# InsertCrossReferencesComfortably
**VBA macro to comfortably insert cross references in MS Word**  

![Screenshot](https://github.com/Traveler4/InsertCrossReferencesComfortably/blob/master/Screenshot.png)
When inserting cross references in MS Word, have you ever been annoyed how tedious it is to do via the miniature default dialog (see screenshot above)? One has to  
- Navigate through the ribbon to find Insert -> Links -> Cross Reference (only visible if the Word window is very wide) 
- Click it
- Select what shall be cross-referenced
- Select how the cross-reference shall be displayed
- Find the element in the list (which, prior to Office 2016 is diminutive and fixed in height !) and click it

How about that instead:
- Press a hotkey when you want to insert a crossreference at the current location
- Put the cursor to the element to which the crossreference shall point to, e.g. a figure label or a headline
- Press the hotkey again => the cursor jumps back to the original location and the crossref is inserted

An additional feature is that you can configure yourself in which form the crossreference is inserted. This allows for combined inserts like the following (text in italics comes from field codes):
  - (see _Figure 7_ _above_, on page _8_) 
  - see _above_, chapter _1_ - _Introduction_
  - cf. Table _2_ (p. _123_)<br>
You can configure multiple formats and toggle between them.
  
## Preparation:  
### Quick way: 
Download the _Example file.docm_ (including macro), open it and play around with it. <br>
However, item 1) from the 'Proper way' is mandatory.
### Proper way:
1) Make sure, the following References are ticked in the VBA editor:<br>
       - Microsoft VBScript Regular Expressions 5.5<br>
    How to do it: https://www.datanumen.com/blogs/add-object-library-reference-vba/
 2) Put the macro code in a VBA module in your document or document template.<br>
    It is recommended to put it into _normal.dot_, 
    then the functionality is available in any document.<br>
 3) Assign a keyboard shortcut to this macro (recommendation: Ctrl+Alt+Q)<br>
    This works like this (Office 2010):
      - File -> Options -> Adapt Ribbon -> Keyboard Shortcuts: Modify...<br>
      - Select from Categories: Macros<br>
      - Select form Macros: [name of the Macro]<br>
      - Assign keyboard shortcut...<br>
 4) Alternatively to 3) or in addition to the shortcut, you can assign this 
    macro to the ribbon button _Insert -> CrossReference_.<br>
    However, then you will not be able any more to access Word's dialog
    for inserting cross references.<br>
      - To assign this sub to the ribbon button _Insert -> CrossReference_,
      just rename this sub to *InsertCrossReference* (without underscore).<br>
      - To de-assign, re-rename it to something like 
      *InsertCrossReference_* (with underscore).
 5) Adapt the configuration according to your preferences.

## Useage:  
1) At the location in the document, where the crossreference shall be inserted,  
   press the keyboard shortcut.  <br>
   A temporary bookmark is inserted (if their display is enabled, grey square brackets will appear).  
2) Move the cursor to the location to where the crossref shall point.  <br>
   Supported are:  
   - Headlines  
   - Subtitles of Figures realised via { SEQ Figure}, e.g. "Figure 123", "Figure 12-345"  
   - Subtitles of Tables  realised via { SEQ Table} , e.g. "Table 123", "Table 12-345"  
   - References to documents realised via { SEQ Ref}, e.g. "[42]"  
   _Recommendation for large documents:_ use the navigation pane (View -> Navigation -> Headlines)   
3) Press the keyboard shortcut again.  <br>
   The cursor will jump back to the location of insertion and the crossref will be inserted. <br>
   Done!  
4) Additional function:  <br>
   Positon the cursor at a cross reference field (if you have configured chained cross reference fields, put the cursor to the last field in the chain).<br> 
   Press the keyboard shortcut.<br>
   - The field display toggles to the next configured option, e.g. from _see Chapter 1_ to _cf. Introduction_.
   - Subsequently added cross references will use the latest format (persistent until Word is exited).  
  
## Limitations:
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
* 190709 Expanded configuration possibilities due to intermediate text sequences
