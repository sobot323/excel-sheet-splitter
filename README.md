# Excel Sheet Splitter

Prosta aplikacja webowa do dzielenia skoroszytów Excel na osobne pliki dla każdego arkusza.

## Funkcje

- Obsługa plików `.xlsx`, `.xls`, `.xlsm` i `.xlsb`.
- Każdy arkusz zapisywany jest jako osobny plik Excel `.xlsx` oraz PDF.
- Wszystkie pliki trafiają do jednej paczki ZIP.
- Nazwy plików składają się z nazwy arkusza i edytowalnego sufiksu.
- Domyślny sufiks ma format `_karta_pracy_MM_YYYY` i jest ustawiany według bieżącej daty.
- Przetwarzanie odbywa się lokalnie w przeglądarce. Plik źródłowy nie jest wysyłany na serwer.

## Uruchomienie

Otwórz [aplikację](https://sobot323.github.io/excel-sheet-splitter/) albo uruchom `index.html` lokalnie w przeglądarce z dostępem do Internetu (biblioteki są ładowane z CDN).

## Pliki PDF

PDF zawiera dane arkusza w postaci tabeli. Aplikacja dobiera orientację strony do liczby kolumn, powtarza pierwszy wiersz jako nagłówek tabeli i automatycznie dzieli długie arkusze na strony.

Eksport PDF nie jest pełnym odwzorowaniem silnika wydruku Excela. Wykresy, makra, obrazy, złożone style, ustawienia wydruku, scalone komórki i inne zaawansowane elementy mogą nie zostać odwzorowane.

## Ograniczenia eksportu Excel

Wynikowe pliki Excel są zapisywane w formacie XLSX. Dane komórek i dostępne wartości formuł są przenoszone, ale konwersja nie gwarantuje zachowania makr, tabel przestawnych, połączeń zewnętrznych ani wszystkich zaawansowanych elementów i stylów skoroszytu. Plik XLSM nie zachowuje makr w wynikowym XLSX.

## Prywatność

Przetwarzanie pliku odbywa się w pamięci przeglądarki. Aplikacja korzysta z bibliotek SheetJS, JSZip i pdfmake pobieranych z CDN, ale nie wysyła zawartości skoroszytu na własny serwer.
