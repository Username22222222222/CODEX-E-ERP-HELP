# Company-Script.sql Review

Quelle: `C:\X-ERP\X-IMPORT\Company-Script.sql`

Gesicherte Arbeitskopie: `C:\Users\micha\Documents\X-ERP-HELP\reference\sql\Company-Script.sql`

## Ergebnis

- Zeilen: 51.291
- Groesse: 2.827.025 Bytes
- Batches: 871
- ScriptDom-Parser: keine Syntaxfehler
- Echte CREATE-Objekte: 1.005
- Doppelte CREATE-Objekte: 0

## Objektuebersicht

- Tabellen: 408
- Views: 95
- Prozeduren: 307
- Funktionen: 4
- Trigger: 7
- Types: 8
- Indexe: 176

## Auffaelligkeit

In `xPROC220D_DocPosition_Delete` steht in der Statusberechnung:

```sql
WHEN QuantityOpen + @QuantityTransferred < Quantity
 AND QuantityOpen - @QuantityTransferred > 0 THEN 'partially'
```

Da die Zeile davor und danach mit `QuantityOpen + @QuantityTransferred` arbeiten, ist das Minus in der zweiten Bedingung fachlich auffaellig. Vermutlich sollte auch dort mit `+ @QuantityTransferred > 0` geprueft werden. Das ist kein Parserfehler, aber ein konkreter Logik-Kandidat fuer eine manuelle fachliche Pruefung.

## Hinweise

- Viele `DROP TABLE`-Treffer betreffen nur temporaere Tabellen (`#...`) und sind beim statischen Review nicht als Fehler zu werten.
- `SELECT *` kommt mehrfach vor. Das ist in Prozeduren und Debug-/Rueckgabeabfragen nicht automatisch falsch, aber bei stabilen API-/UI-Rueckgaben wartungsanfaelliger als explizite Spaltenlisten.
- Die SQLCMD-Parse-Pruefung gegen `MICRO\X` konnte wegen ODBC-/Verschluesselungsverhandlung nicht aufgebaut werden. Die lokale ScriptDom-Pruefung war erfolgreich.
