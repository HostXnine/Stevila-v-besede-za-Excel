# Števila v besede za Excel
### Prenesi modul [SpellNumber.bas](https://github.com/HostXnine/Stevila-v-besede-za-Excel/releases/download/v1.0.0/SpellNumber.bas) in ga uvozi v Excel.

Pretvarjanje števil v besede za Excel VBA.

Primer: število 1200,22 zapiše kot "tisoč dvesto 22/100".

Modul za Excel je predvsem namenjen za uporabo pri spajanju dokumentov kjer je potrebno npr. v pogodbe zapisati zneske z besedo.  Res je da obstajajo rešitve za pretvorbo števil v Wordu, toda v določenih primerih Word narobe zapiše števila.

Dodajte modul v Excel ([lahko si pomagate s temi navodili](https://support.microsoft.com/sl-si/office/pretvarjanje-%C5%A1tevil-v-besede-a0d166fb-e1ea-4090-95c8-69442cd55d98)) ali prenesete Excelovo datoteko [SpellNumber.xlsm](https://github.com/HostXnine/Stevila-v-besede-za-Excel/releases/download/v1.0.0/SpellNumber.xlsm) v katerem je že dodan modul toda v tem primeru fukcija ne bo delovala, ker jo Excel samodejno blokira zaradi varnostnih razlogov.

### Navodila za uvoz:
1. Odprite Excelov zvezek
2. Pritisnite Alt + F11 da se odpre urejevalnik Visual Basic for Aplications (VBA)
3. Izberite File > Import File v pojavnem oknu izberite preneseno datoteko SpellNumbers.bas.
4. Pritisnete shrani v pojavnem oknu pod shrani kot izberite Excel Macro-Enabled Workbook (.xlsm) in pritisnete shrani
5. Zaprite urejevalnik VBA v kolikor se že ni sam zaprl.

### Navodila za zporabo funkcije:
Uporaba funkcije SpellNumber v posameznih celicah
1. Vnesite formulo **=SpellNumber(A1)**, v celico, kjer želite prikazati napisanih številko, kjer je A1 celice, ki vsebujejo števila ga želite pretvoriti. Lahko tudi ročno vnesite želeno vrednost na primer = SpellNumber(22.50).

2. Pritisnite Enter , da potrdite formulo.

### Viri: 
Vse zasluge za izvirno kodo gredo njihovim avtorjem na spodnjih povezavah: 
https://stackoverflow.com/questions/51204004/convert-numbers-to-words-with-vba
https://support.microsoft.com/sl-si/office/pretvarjanje-%C5%A1tevil-v-besede-a0d166fb-e1ea-4090-95c8-69442cd55d98
