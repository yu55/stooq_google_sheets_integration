# stooq_google_sheets_integration.gs

![stooq_integration.png](screenshots/stooq_integration.png?raw=true "Arkusz pobierający dane z portalu Stooq.com")

`stooq_google_sheets_integration.gs` to skrypt który w bardzo podstawowy sposób "integruje" portal [stooq.com](https://stooq.com) z aplikacją Google Arkusze. Skrypt wprowadza nową funkcję która pobiera kurs wybranego waloru ze strony [stooq.com](https://stooq.com) i umieszcza ten kurs w naszym arkuszu kalkulacyjnym.

## Instrukcja

1. Utwórz swój nowy arkusz kalkulacyjny w aplikacji "Google Arkusze"

2. Dodaj do nowo utworzonego arkusza skrypt `stooq_google_sheets_integration.gs`. Aby to zrobić kliknij w swoim arkuszu "Narzędzia -> Edytor skryptów" i otwarta zostanie nowa zakładka przeglądarki z zawartością edytora skryptów. Usuń treść skryptu w edytorze (jeśli jest jakaś domyślna) i wklej do edytora CAŁĄ (łącznie z komentarzami) zawartość pliku [`stooq_google_sheets_integration.gs`](stooq_google_sheets_integration.gs). Zapisz skrypt wykonując: "Plik -> Zapisz". Jeśli aplikacja zapyta o nazwę projektu to możesz podać "stooq_google_sheets_integration.gs".

![stooq_script_editor.png](screenshots/stooq_script_editor.png?raw=true "Widok edytora skryptów")

3. Zamknij zakładkę przeglądarki z edytorem skryptów.

4. Zamknij i ponownie otwórz zakładkę przeglądarki z arkuszem kalkulacyjnym. Po otwarciu pojawi się nowe menu "STOOQ" u góry arkusza.

5. W swoim arkuszu obowiązkowo musisz zarezerwować jedną komórkę na "datę notowań" dla skryptu `stooq_google_sheets_integration.gs`. Bez tej komórki skrypt nie może prawidłowo funkcjonować. Domyślnie jest to komórka "B1", ale możesz ją zmienić w skrypcie (linijka 25 w treści skryptu).

6. Teraz możesz zacząć używać w swoim arkuszu nowej funkcji którą dostarcza skrypt:

`STOOQ_GET_PRICE("WIG20"; $B$1)` - funkcja zwróci ostatnią wartość kursu indeksu "WIG20" (pobierze ją wprost ze strony https://stooq.com/q/?s=wig20). Zamiast "WIG20" możesz podać dowolny inny symbol ze Stooq, np. "^SPX" czy "BTCUSD". Funkcja musi korzystać z daty w komórce `$B$1`. Naciśnięcie przycisku w menu "STOOQ -> Aktualizuj kursy walorów" spowoduje zaktualizowanie daty w komórce `B1` i pobranie przez wszystkie funkcje `STOOQ_GET_PRICE` najnowszych kursów.
Można w komórce `B1` również ręcznie wpisać datę z przeszłości co spowoduje pobranie historycznego kursu zamknięcia w danym dniu.

Komórka daty `B1` zawiera oprócz daty również czas. Czas jest ignorowany przez funkcję `STOOQ_GET_PRICE` ale niestety gdy będzie tam tylko sama data to Google Arkusze potraktują to jako argument funkcji `STOOQ_GET_PRICE` który się nie zmienia i będą zwracać zawsze tę samą wartość tej funkcji, to jest wartość uzyskaną podczas pierwszego uruchomienia.
Funkcja `STOOQ_GET_PRICE` uruchamia się zawsze automatycznie zaraz po otwarciu arkusza kalkulacyjnego.


