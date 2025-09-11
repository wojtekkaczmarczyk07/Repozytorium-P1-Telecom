Instrukcja (wersja v2 — shadow DOM + dopasowanie po Y):

1) Rozpakuj folder i załaduj jako „Load unpacked” w chrome://extensions (Developer mode).
2) Wejdź na https://crm.playmobile.pl/ (lub miejsce z listą).
3) Ctrl+A, Ctrl+C — rozszerzenie nie blokuje kopiowania, ale po nim nadpisuje tekst w schowku:
   dopina „ DB” po NIP-ach, dla których w tym samym wierszu na ekranie widać „DB” (również jeśli „DB” jest w shadow DOM lub jako pseudo-element).
4) Z perspektywy strony WWW jest to zwykłe kopiowanie (zmiana dzieje się w Twoim schowku).

Jeśli „DB” nadal nie trafia:
- Zwiększ ROW_Y_TOL w content.js (np. do 30–36), jeśli wiersze są wysokie.
- Upewnij się, że w chwili kopiowania „DB” jest widoczne w tym samym wierszu (nie zwinięte, nie poza viewportem).