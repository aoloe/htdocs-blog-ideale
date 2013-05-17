I've been asked to create an Excel sheet for managing the kids going to a Kindergarden... I don't know why, but I've said:ok, let's do it.

Here are some notes about my experience! (if there is any interest for it, I can translate the macros from German to English)

- There is no Sum(Above) in Excel (but there is one in Word...) but there is an ugly workaround:
      =SUMME(INDIREKT(ADRESSE(1;SPALTE())):INDIREKT(ADRESSE(ZEILE()-1;SPALTE())))
  Wow!

- If you want to have a list of values (strings), just add the values (to a separate sheet and then use
      Daten > Datenüberprüfung > Datenüberprüfung... > Zulassen = Liste & Quelle = [select the cells]

- If you want to create some macros, you have to first right click on the toolbar, choose "Menüband anpasen" and in the "Hauptregister" list activate the "Entwickler" option.
