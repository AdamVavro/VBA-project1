# VBA aplikácia pre optimalizáciu práce s plánmi upínania

Tento projekt vznikol pri snahe zmodernizovať a zefektívniť zastarelé procesy tvorby a aktualizácie plánov upínania pre lisovacie nástroje.

<!--## Popis projektu-->
Keď som začal pracovať na tomto projekte nevedel som o programovaní nič :exploding_head: . Ani to vlastne nebol pôvodny zámer vyrobiť niečo, kde budem využívať programovanie. Postupne, pri riešení problémov som však začal prenikať do sveta programovania a chytilo ma to natoľko, že som sa začal v tom vzdelávať aj vo voľnom čase.  

Aplikácia je výsledkom mojej 1,5 ročnej práce, počas ktorej som 9 mesiacov vyvýjal a 9 mesiacov testoval a dolaďoval vo výrobe.  
Nie je to najvyšší level programovania, mnohé veci by som už teraz vedel spraviť jednoduchšie, avšak aplikácia funguje a pre potreby firmy plní svoj účel, tak už do nej nezasahujem...

## Stručný popis projektu
Aplikácia slúži na jednoduché vytvorenie plánu upínania (ďalej len "plán") aj pre menej zdatného používateľ. Formulár ho navedie na doplnenie všetkých potrebných údajov a následné po kliknutí na tlačidlo "Uložiť plán upínania" sa automaticky uloží na všetkých potrebných miestach a v potrebných formátoch. Po uložení sa údaje exportujú do databázy, odkiaľ sa v prípade potreby dajú automaticky načítať do plánu.  

### Skladá sa z troch častí:  
- [Formulár](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/06.%20Formul%C3%A1r.jpg)
- [Plán upínania](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/01.%20Pr%C3%A1zdny%20pl%C3%A1n%20up%C3%ADnania.jpg)
- [Databáza](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/27.2%20Datab%C3%A1za.jpg)

![Alternatívny text](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/00.%20Komplet.jpg)


<!--1.TEST__________________________________________________________________________________________________________________-->

<!--<details><summary>1.<ins>TEST</ins></summary>
	
![SCRENSHOT](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/01.%20Pr%C3%A1zdny%20pl%C3%A1n%20up%C3%ADnania.jpg)
<details><summary>kód</summary>

![CODE](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Code/Code%20screenshots/01.%20Po%20otvoren%C3%AD%20nastav%C3%AD%20ve%C4%BEkos%C5%A5%20okna.jpg)
</details>
 
---
</details>-->
<!--1._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>1.<ins>Po otvorení nastaví veľkosť a polohu okna, skryje riadok vzorcov, skryje záhlavia, skryje mriežku a zbalí pás s nástrojmi.</ins></summary>

![SCRENSHOT](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/01.%20Pr%C3%A1zdny%20pl%C3%A1n%20up%C3%ADnania.jpg)
<details><summary>kód</summary>

![CODE](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Code/Code%20screenshots/01.%20Po%20otvoren%C3%AD%20nastav%C3%AD%20ve%C4%BEkos%C5%A5%20okna.jpg)
</details>

---
</details>
<!--1.	[Po otvorení nastaví veľkosť a polohu okna, skryje riadok vzorcov, skryje záhlavia, skryje mriežku a zbalí pás s nástrojmi.](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/01.%20Pr%C3%A1zdny%20pl%C3%A1n%20up%C3%ADnania.jpg)<details><summary>Kód</summary>![Alternatívny text](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Code/Code%20screenshots/01.%20Po%20otvoren%C3%AD%20nastav%C3%AD%20ve%C4%BEkos%C5%A5%20okna.jpg)</details>-->

<!--2._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>2.<ins>Po zadaní čísla nástroja sa zobrazí dialógové okno, ktoré sa spýta či si prajete doplniť základné údaje o nástroji z databázy.</ins></summary>

![SCRENSHOT](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/02.%20Doplni%C5%A5%20%C3%BAdaje.jpg)
<details><summary>kód</summary>

![CODE](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Code/Code%20screenshots/02.%20%C4%8C%C3%ADslo%20n%C3%A1stroja.jpg)
</details>

---
</details>	 
<!--2.	[Po zadaní čísla nástroja sa zobrazí dialógové okno, ktoré sa spýta či si prajete doplniť základné údaje o nástroji z databázy.](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/02.%20Doplni%C5%A5%20%C3%BAdaje.jpg)<details><summary>Kód</summary>![Alternatívny text](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Code/Code%20screenshots/02.%20%C4%8C%C3%ADslo%20n%C3%A1stroja.jpg)</details>-->

<!--3._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>3.<ins>Po potvrdení otvorí databázu, vyhľadá číslo nástroja, načíta príslušné informácie.</ins></summary>

![SCRENSHOT](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/03.%20Na%C4%8D%C3%ADtanie%20%C3%BAdajov%20z%20datab%C3%A1zy.jpg)
<details><summary>kód</summary>

![CODE](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Code/Code%20screenshots/03.-04.%20Na%C4%8D%C3%ADta%20%C3%BAdaje%20z%20datab%C3%A1zy%20a%20dopln%C3%AD%20do%20pl%C3%A1nu.jpg)
</details>

---
</details>

<!--3.	[Po potvrdení otvorí databázu, vyhľadá číslo nástroja, načíta príslušné informácie.](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/03.%20Na%C4%8D%C3%ADtanie%20%C3%BAdajov%20z%20datab%C3%A1zy.jpg)<details><summary>Kód</summary>![Alternatívny text](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Code/Code%20screenshots/03.-04.%20Na%C4%8D%C3%ADta%20%C3%BAdaje%20z%20datab%C3%A1zy%20a%20dopln%C3%AD%20do%20pl%C3%A1nu.jpg)</details>-->

<!--4._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>4.<ins>Údaje sa doplnia do plánu.</ins></summary>

![SCRENSHOT](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/04.%20Automatick%C3%A9%20doplnenie%20%C3%BAdajov.jpg)
<details><summary>kód</summary>

![CODE](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Code/Code%20screenshots/03.-04.%20Na%C4%8D%C3%ADta%20%C3%BAdaje%20z%20datab%C3%A1zy%20a%20dopln%C3%AD%20do%20pl%C3%A1nu.jpg)
</details>

---
</details>

<!--4.	[Údaje sa doplnia do plánu.](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/04.%20Automatick%C3%A9%20doplnenie%20%C3%BAdajov.jpg)<details><summary>Kód</summary>![Alternatívny text](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Code/Code%20screenshots/03.-04.%20Na%C4%8D%C3%ADta%20%C3%BAdaje%20z%20datab%C3%A1zy%20a%20dopln%C3%AD%20do%20pl%C3%A1nu.jpg)</details>-->

<!--5._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>5.<ins></ins></summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
</details>


<!--5.	[Klikneme na tlačidlo „Otvoriť formulár“. Okno s plánom sa minimalizuje a v ľavej časti obrazovky sa zobrazí formulár pre zápis údajov z modelu.](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/05.%20Tla%C4%8Didlo%20otvori%C5%A5%20formul%C3%A1r.jpg)<details><summary>Kód</summary>![Alternatívny text](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Code/Code%20screenshots/05.%20Tla%C4%8Didlo%20Otvori%C5%A5%20formul%C3%A1r.jpg)</details>-->

<!--6._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>6.<ins></ins></summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
</details>


<!--6.	[Formulár.]()<details><summary>Kód</summary>![Alternatívny text]()</details>-->

<!--7._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>7.<ins>Otvoríme 3D model nástroja a pozapisujeme všetky údaje do formulára.</ins></summary>

<!--![SCRENSHOT](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/07.%20Otvorenie%20CAD%20modelu.jpg)-->

<!--_________________________________________________________7.1_______________________________________________________________________________________________________________________________________-->
<details><summary>Rozmery nástroja D, Š, V.</summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
![CODE]()
</details>

---
</details>


	[Otvoríme 3D model nástroja a pozapisujeme všetky údaje do formulára.](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/07.%20Otvorenie%20CAD%20modelu.jpg)
    <details>
      
      <summary>
        
      

  	  </summary>
     
      -	[Rozmery nástroja D, Š, V.]()<details><summary>Kód</summary>![Alternatívny text]()</details>

      -	[Vzdialenosť medzi drážkami]()<details><summary>Kód</summary>![Alternatívny text]()</details>

      -	[GDF(OB)]()<details><summary>Kód</summary>![Alternatívny text]()</details>

      -	[Upínacia výška nástroja]()<details><summary>Kód</summary>![Alternatívny text]()</details>

      -	[Prítomnosť pridržiavača alebo GDF a možnosť upnutia do lisov PWS.]()<details><summary>Kód</summary>![Alternatívny text]()</details>

      -	[Dialógové okno „Prajete si vyznačiť pozíciu tlačných čapov?“.]()<details><summary>Kód</summary>![Alternatívny text]()</details>

      -	[Po potvrdení sa formulár zatvorí okno s plánom zmení rozmer a presunie sa vľavo dole a zobrazí raster stola. Po aktivovaní bunky v rastri stola sa zobrazia tlačidlá „Centrovanie“, „Tlačný čap“, „Voľné miesto“, „OK“.]()<details><summary>Kód</summary>!     [Alternatívny text]()</details>

      -	[Pomocou zobrazených tlačidiel sa vyznačia pozície tlačných čapov.]()<details><summary>Kód</summary>![Alternatívny text]()</details>

      -	[Keď je všetko vyznačené pomocou tlačidla „OK“ sa okno zavrie a opäť sa zobrazí formulár.]()<details><summary>Kód</summary>![Alternatívny text]()</details>

      -	[Priemer centrovania]()<details><summary>Kód</summary>![Alternatívny text]()</details>

      -	[Po zapísaní súradníc centrovania z modelu sa automaticky prevedú na súradnice plánu upínania  a podľa nich sa vyznačí v rastri stola pozícia centrovania.]()<details><summary>Kód</summary>![Alternatívny text]()</details>

      -	[Po zadaní smeru lisovania sa v pláne zobrazí smer lisovania]()<details><summary>Kód</summary>![Alternatívny text]()</details>

      - [Vyplniť poznámky]()<details><summary>Kód</summary>![Alternatívny text]()</details>
      
     </details>

<!--8._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>8.<ins></ins></summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
</details>


 <!-- 8. <details><summary>Po kliknutí na tlačidlo "Zatvoriť formulár" sa formulár zavrie.</summary>
![Alternatívny text](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Screenshots/%C5%A4ahovka/00.%20Komplet.jpg)
        ![Alternatívny text](https://github.com/AdamVavro/VBA-project1/blob/KT05_05/Code/Code%20screenshots/20.%20Pozn%C3%A1mky%2C%20Zatvorit%20formul%C3%A1r.jpg)</details>-->

 <!--9._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>9.<ins></ins></summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
</details>


  <!--9.	[Zobrazí sa plán s doplnenými údajmi.]()<details><summary>Kód</summary>![Alternatívny text]()</details>-->

  <!--10._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>10.<ins></ins></summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
</details>


 <!-- 10.	[Po aktivovaní bunky v oblasti poznámky k nástroju sa zobrazia tlačidlá pre formátovanie poznámok.]()<details><summary>Kód</summary>![Alternatívny text]()</details>-->

<!--11._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>11.<ins></ins></summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
</details>

  

 <!-- 11.	[Po zapísaní údajov do oblasti „Pracovné tlaky a nastavenia“ sa zobrazí dialógové okno, ktoré sa spýta, či chcem údaj zapísať do databázy.]()<details><summary>Kód</summary>![Alternatívny text]()</details>-->

<!--12._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>12.<ins></ins></summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
</details>


 <!-- 12.	[Po potvrdení sa otvorí databáza vyhľadá sa číslo nástroja a konkrétna bunka s príslušným údajom. Údaj sa zapíše do bunky v databáz a v komentári bunky ktorý slúži ako archív sa zaznamená dátum, čas a zapísaná hodnota.]()<details><summary>Kód</summary>![Alternatívny text]()</details>-->

<!--13._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>13.<ins></ins></summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
</details>


 <!-- 13.	[Po kliknutí na tlačidlo „Uložiť plán upínania“ sa zobrazí dialógové okno, kde sa zobrazí názov uloženého súboru a adresa kde sa plán uloží.]()<details><summary>Kód</summary>![Alternatívny text]()</details>-->

<!--14._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>14.<ins></ins></summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
</details>


<!--  14.	[Po potvrdení sa plán uloží 3x ako pdf súbor, 1x ako xlsm súbor a 1x ako jpg súbor.]()<details><summary>Kód</summary>![Alternatívny text]()</details>-->

<!--15._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>15.<ins></ins></summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
</details>


 <!-- 15.	[Všetky údaje z plánu upínania sa exportujú do databázy]()<details><summary>Kód</summary>![Alternatívny text]()</details>-->

<!--16._________________________________________________________________________________________________________________________________________________________________________________________________-->
<details><summary>16.<ins></ins></summary>

![SCRENSHOT]()
<details><summary>kód</summary>

![CODE]()
</details>

---
</details>

 <!-- 16.	[V prípade potreby sa dajú do plánu importovať údaje z databázy naspať.]()<details><summary>Kód</summary>![Alternatívny text]()</details>-->



