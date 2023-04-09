# -stevila-v-besede-excel
Pretvarjanje števil v besede za Excel VBA.

Program je predvsem namenjen za uporabo spajanje dokumentov kjer je potrebno v pogodbe zapisati zneske z besedo. Na primer število 1200,22 zapiše kot "tisoč dvesto 22/100"

Ko naložite program v Excel ([lahko si pomagate s temi navodili](https://support.microsoft.com/sl-si/office/pretvarjanje-%C5%A1tevil-v-besede-a0d166fb-e1ea-4090-95c8-69442cd55d98)) uporabite formulo =SpellNumber().

### Podrobnejša navodila:
Uporaba funkcije SpellNumber v posameznih celicah
1. Vnesite formulo = SpellNumber(A1), v celico, kjer želite prikazati napisanih številko, kjer je A1 celice, ki vsebujejo števila ga želite pretvoriti. Lahko tudi ročno vnesite želeno vrednost na primer = SpellNumber(22.50).

2. Pritisnite Enter , da potrdite formulo.

### Viri: 
Vse zasluge za izvirno kodo gredo njihovim avtorjem na spodnjih povezavah: 
https://stackoverflow.com/questions/51204004/convert-numbers-to-words-with-vba
https://support.microsoft.com/sl-si/office/pretvarjanje-%C5%A1tevil-v-besede-a0d166fb-e1ea-4090-95c8-69442cd55d98
