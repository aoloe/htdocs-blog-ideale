I've been asked to create an Excel sheet for managing the kids going to a Kindergarden... I don't know why, but I've said:ok, let's do it.

Here are some notes about my experience! (if there is any interest for it, I can translate the macros from German to English)

- There is no Sum(Above) in Excel (but there is one in Word...) but there is an ugly workaround:
      =SUMME(INDIREKT(ADRESSE(1;SPALTE())):INDIREKT(ADRESSE(ZEILE()-1;SPALTE())))
  Wow!

- If you want to have a list of values (strings), just add the values (to a separate sheet and then use
      Daten > Datenüberprüfung > Datenüberprüfung... > Zulassen = Liste & Quelle = [select the cells]

- And if you want to sum the values corresponding to a dropdown list:
      =SUMME(WENN(INDIREKT(ADRESSE(1;SPALTE())):INDIREKT(ADRESSE(ZEILE()-1;SPALTE()))="am";0.7))+
       SUMME(WENN(INDIREKT(ADRESSE(1;SPALTE())):INDIREKT(ADRESSE(ZEILE()-1;SPALTE()))="pm";0.5))
  This is a "Matrice" operation, so you will have to press ctrl+shift+enter to quit the formula editing.

- If you want to create some macros, you have to first right click on the toolbar, choose "Menüband anpasen" and in the "Hauptregister" list activate the "Entwicklertools" option.

- Here is the macro i've created to copy the currently selected five cells n times:
      Sub multiplikation()
      '
      ' multiplikation Makro
      ' 5 ausgewählten zellen n mal nach rechts kopieren
      '
      '
      Dim n
          ' Debug.Print
          If (Not Selection.Columns.Count = 5) And (Not Selection.Rows.Count = 5) Then
              MsgBox ("Please select one week of five days")
              Exit Sub
          End If
          n = InputBox("Anzahl Wochen", "Wochen duplizieren")
          If (Not IsNumeric(n)) Or (Not n > 0) Then
              MsgBox ("No week entered")
              Exit Sub
          End If
          Selection.Copy
          r = Selection.Row()
          c = Selection.Column()
          For i = 1 To n
              Cells(r, c + (i * 5)).Select
              ActiveSheet.Paste
          Next i
          
      End Sub

- Adding a button:
  "On the Developer tab, in the Controls group, click Insert, and then under Form Controls, click Button"
  Then assign the macro through the context menu on the button.

- If you want to avoid that during the user can edit the field formulas, in "start > cells > format > Blatt schützen" activate the option "Formattieren" (in the same dialog you can then switch on and the off the option)

- If you want to keep some columns always visible on the left, select the first "normal" column and run "Ansicht > Fenster einfrieren > Fenster efinfrieren"

- If you want to print some columns and have the first n coloumns printed on the left of them, go to "Seitenlayout > Blattoptionen" and in the "Blatt" tab set the "Wiederholungsspalten links" to the columns you have fixated above.
