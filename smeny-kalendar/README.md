# smeny@Praha.vbs

Tento skript slouží k převodu dat z Excelu (`Export.xlsx`) do CSV souboru (`Final.csv`) pro různé lokality (Praha, Chrudim, Ostrava, Homeoffice). Výstupní CSV je vhodný například pro import do kalendáře.

## Jak skript funguje

1. **Výběr lokality:**  
   Lokalita je určena podle názvu skriptu (např. `smeny@Praha.vbs`). Podle lokality se nastaví adresa pro výstupní soubor.

2. **Načtení dat:**  
   Skript otevře Excel soubor `Export.xlsx` a načte list `Sheet1`.

3. **Zpracování řádků:**  
   Pro každý řádek od třetího řádku:
   - Pokud ve sloupci 5 není "Yes":
     - Pokud je ve sloupci 18 hodnota `NVP-M` **nebo** je ve sloupci 9 cokoliv, zapíše se řádek s typem `Volno`.
     - Pokud je ve sloupci 13 cokoliv, zapíše se `Dovolená`.
     - Pokud je ve sloupci 11 cokoliv, zapíše se `SickDay`.
     - Pokud je ve sloupci 7 cokoliv, zapíší se tři řádky: `Směna`, `Oběd`, `Směna`.
     - Jinak se zapíše pouze jedna směna.

4. **Výstup:**  
   Výsledný CSV soubor je uložen jako `Final.csv` v UTF-8.

## Požadavky

- Windows s podporou VBScript
- Microsoft Excel (pro čtení .xlsx souboru)
- Soubor `Export.xlsx` ve stejném adresáři jako skript

## Použití

1. Upravte název skriptu podle požadované lokality, např. `smeny@Praha.vbs`, `smeny@Chrudim.vbs`, atd.
2. Umístěte skript a soubor `Export.xlsx` do stejné složky.
3. Spusťte skript dvojklikem nebo přes příkazovou řádku:
   ```
   cscript smeny@Praha.vbs
   ```
4. Výsledný soubor `Final.csv` najdete ve stejné složce.

## Poznámky

- Pokud není nalezena lokalita v názvu skriptu, skript se ukončí s chybovou hláškou.
- Pokud není nalezen vstupní Excel nebo list, skript se ukončí s chybovou hláškou.
- Výstupní CSV je vždy přepsán.

---
Autor:  
GitHub Copilot