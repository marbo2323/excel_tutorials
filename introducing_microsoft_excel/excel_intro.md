
# Micorosoft Exceli tutvustus

## Miks kasutada Microsoft Excelit
Eksisteerib mitmeid tarkvara lahendusi arvutustabelitega (ik. *spreadsheet*) töötamiseks. 
Google Sheets on paljude funktsioonidega toimiv tasuta lahendus, millel on paljud võimalused nagu ka kommerts töölaua rakendustel.
Siiski, kui planeerite arvutustabelite kasutamist ärilistel eesmärkidel, kasutate suure tõenäosusega [Microsofti excel'it](https://www.microsoft.com/en-us/microsoft-365/excel).
See on kõige populaarsem tarkvara arvutustabelite loomiseks. Seda kasutab enamik ettevõtteid andmeanalüüsiks ja visualiseerimiseks.
Excel pole mitte ainult populaarne, vaid ka väga võimas. See võimaldab talletada suuremas mahus andmeid kui vabavaralised alternatiivid ja sisaldab palju äri- ja andmeanalüüsi funktsioone. Excel on tasuline tarkvara. Mõnikord on Excel paigaldatud ka uutesse arvutitesse. Excel on saadaval mitmete operatsioonisüsteemide jaoks nagu Windows, IOS, android. Excelit on võimalik kasutada ka online versioonina.

## Exceli kasutamine Windowsi arvutis

Enamik inimesi kasutab Excelit töölauaversioonina arvutis. Teeme lühiülevaate Exceli kasutajaliidesest. Selline näeb välja tühi Exceli arvutusleht (ik. *sheet*). 
![Tühi Exceli tööleht](/img/empty_excel_workbook.PNG)

Näeme siin tühja tabelit mis koosneb tupladest (ik. *column*)

![Märgistatud tupl](/img/excel_selected_column.PNG) 

ja ridadest (ik. *row*)

![Märgistatud rida](/img/excel_selected%20row.PNG)

Lehe allservas näeme töölehtede vahelehti (ik. *tab*):

![Töölehe vaheleht](/img/excel_sheet_tab.PNG)

Klõpsates "+" märgile, saab lisada uue töölehe. Alternatiivina saab uue töölehe lisada ka menüü valikust "Avaleht" > "Lisa" > "Lisa leht".

![Töölehe lisamine menüüst](/img/excel_add_sheet_menu_command.PNG)

Kolmanda võimalusena saab töölehe lisamiseks kasutada ka klahvikombinatsiooni shift + F11.

Akna ülaosas on kiirpääsu riba (ik. *quick access toolbar*).

![Kiirpääsu riba](/img/excel_quick_access_toolbar.PNG)

Vaikimisi on seal käsud Salvesta (ik. *save*), Võta tagasi (ik. *undo*) ja Tee uuesti (ik. *redo*).

Klõpsates nool alla ikoonil saab kiirpääsu riba kohandada, lisada või eemaldada sealt käskusid.

![Kiirpääsu riba kohandamine](/img/excel_customize_quickaccess_toolbar.PNG)

Järgmisena näeme Exceli tööriistade linti (ik. *ribbon*).

![Lint](/img/excel_ribbon.PNG)

Selle ülemises osas asuvad vahelehed. Igale vahelehele vastab oma tööriistade komplekt. Töölehe avanedes on vaikimis valitud vaheleht nimega Avaleht (ik. *Home*). Kui valida näiteks vaheleht Andmed (ik. *Data*), siis näeme selle vahelehe tööriistu.

![Lint vaheleht Andmed](/img/excel_ribbon_data_toolbar.PNG)

Sellel tööriistaribal on näiteks sellised käsud nagu Sordi (ik. *Sort*) ja Filtreeri (ik. *Filter*). Need võimaldavad muuta andmete järjekorda või filtreerida andmed mingi tunnuse alusel.

Vahelehel Vaade (ik. *Viev*) on näiteks käsud Külmuta paanid (ik. *freeze panes*), küljenduse valikud (ik. *page layout*) ja muud valikud vaate kohandamiseks.

![Lint vaheleht Vaade](/img/excel_ribbon_view_toolbar.PNG)

Lindile järgneb Valemiriba.

![Valemiriba](/img/excel_formula_bar.PNG)

Escelis on väga suur hulk võimalusi ja käske. Seetõttu on ka palju võimalusi tarkvara kohandamiseks enda vajaduste tarbeks. Vaatame hästi lühidalt mõnda kohandamisvõimalust. Eespool oleval Exceli pildil on näha ka lindil Arendaja vaheleht. Vaikimisi seda ei ole. Valides menüüst Fail > Suvandid ja edasi Lindi kohandamine (ik. *File > Options > Customize ribbon*), saame valida millised lindi vahelehed on nähtavad.

![Suvandid lindi kohandamine](/img/excel_options_customize_ribbon.PNG)

Siin saame linnukesega märgistada, milliseid tööriistaribasid soovime näha. Näiteks kui eemaldada linnukle Arendaja tööriistaribalt, ja klõpsata Ok, siis kaob see tööriistariba.

![Lint Arendaja tööristariba eemaldatud](/img/excel_developer_toolbar_hidden.PNG)

Saab luua ka enda kohandatud menüü, ja valida sinna komplekti vajalikke käske.

suvandite aknas on veel väga palju võimalusi. Näiteks saab kohandada rakenduse värvi.

![suvandid muuda Office värv](/img/excel_options_set_office_color.PNG)

Näiteks kui valisime värviks Must:

![Office värv must](/img/excel_office_color_black.PNG)

Excelil on jõuline (ik. *robust*) lühivalikute (ik. *shortcut*) tugi. enamikel käskudel on klaviatuuri lühivalik. Näiteks kui klõpsata klaviatuuril alt klahvi ilmub ekraanile hulk tähti.

![Alt klahvi vajutamine](/img/excel_alt_pressed.PNG)

Need on klaviatuuri lühivalikud. Proovime oma muudetud väri teem atagasi endiseks muuta kasutades ainult klaviatuuri lühivalikuid.

Klõpsame Alt, klõpsame F, et avada Fail menüü.

![Lühivalik Fail menüüsse](/img/excel_shortcut_to_file_menu.PNG)

Suvandite akna avamiseks klõpsame nüüd D. Avaneb eelpool pildidl kuvatud suvandite aken. suvandite aknas saame erinevate valikute vahel liikuda Tab klahvi ja nooleklahvidega. Liigume nüüd valikule Office kujundus. suvandite aknas näeme, et suvandite nimedel on osad tähed allajoonitud.

![Lühivaliku tähisus allajoonimisega](/img/excel_shorcut_underine_marker.PNG)

Klõpsame nüüd alt + c. Sellega avatakse Office kujunduse suvandi valikute loetelu. Liigume nooleklahviga valikule Kasuta süsteemisätet. Tab klahviga saame edasi liikuda suvandite akna Ok nupu peale ja klõpsame Enter. suvandite aken sulgub ja Exceli kujundus muutub.

Kui kasutate igapäevaselt Excelit palju, siis on mõistlik lühibvalikud selgeks õppida. See teeb töö palju kiiremaks. Näiteks kirjutame ühte lahtrisse väga lihtsa valemi.

![Lihtne liitmisvalem](/img/excel_simple_formula.PNG)

Meie eesmärk on kopeerida valemi väärtus teise lahrrisse käsuga Kleebi teisiti: Väärtus (ik. *paste special: value*).

Selleks kopeerime lahtri klahcikombinatsiooniga Ctrl + c, liigume nooleklahviga teise lahtrisse ja kleebime väärtuse klahvikombinatsiooniga Alt + H + V + V (3X).

![Kleebi teisiti: väärtused](/img/excel_paste_special_values.PNG)

Kuna kleebi teisiti valikus on V tähega lühivalikuid mitu, tuleb klõpsata V tähel kuni see jõuab soovitud valikuni (Kleebi väärtused). Seejärel klõpsa Enter.

![Kleebitud väärtus](/img/excel_pasted_value.PNG)

Näeme, et lahtrisse C3 kleebiti valemi asemel selle tulemus 8.
