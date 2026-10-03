# Cenová mapa nájemného - data v JSON

Strojově čitelná kopie **cenové mapy nájemného Ministerstva financí ČR**
(zákon č. 151/1997 Sb., vyhláška č. 456/2024 Sb.). MF ji vydává čtvrtletně jako XLSX;
tenhle repozitář ji automaticky převádí na JSON, který si webové kalkulačky
načtou přímo z CDN.

- Zdroj: [mf.gov.cz - cenová mapa nájemného](https://mf.gov.cz/cs/rozpoctova-politika/podpora-projektoveho-rizeni/cenova-mapa/cenova-mapa-infografika)
- Používá: [investovanihrou.cz](https://investovanihrou.cz/kalkulacky/)

## Soubory

| Soubor | Obsah |
|---|---|
| `data/index.json` | `meta` (zdroj, období, výměry, příplatky za vybavení) + `lokality`: `[katastr, obec, kód kraje, pořadí v souboru kraje]` |
| `data/kraj-NN.json` | `[katastr, obec, kód obce, {VK: [ref, dolní, horní, novostavba, min, max, medián, pokrytí]}]` |

Hodnoty jsou v Kč za m² a měsíc, bez služeb. VK = velikostní kategorie
(1 = 1+kk/1+1, 2 = 2+kk/2+1, 3 = 3+kk/3+1, 4 = 4+). Pokrytí 1-5 = počet
pozorování v lokalitě (1 = 0-30 inzerátů, 5 = 150+).

CDN: `https://cdn.jsdelivr.net/gh/miloshamsa/cenova-mapa-najemneho@main/data/index.json`

## Co data znamenají

MF počítá **nabídkové** ceny z realitních portálů za posledních 5 let (vážené
k poslednímu roku) pro modelový byt: nezařízený, starší 5 let, bez balkonu,
výtahu a garáže. Je to proto **konzervativní odhad**: srovnání s 9 608 aktuálními
inzeráty na Sreality (3. 10. 2026) ukázalo, že dnešní nabídka je v mediánu
o 22 % výš (polovina inzerátů +4 až +42 %).

## Aktualizace

GitHub Action `cenova-mapa.yml` běží 16. a 25. února, května, srpna a listopadu
(MF vydává do 45 dnů po čtvrtletí). Nový soubor commitne a vyprázdní cache jsDelivr.

Když MF změní strukturu souboru, `scripts/cenova_mapa.py` skončí chybou, nic
nepřepíše a běh Action spadne (přijde e-mail). Ruční běh:

```bash
pip install -r scripts/requirements.txt
python scripts/cenova_mapa.py                    # nejnovější soubor z webu MF
python scripts/cenova_mapa.py --soubor mapa.xlsx # lokální soubor
```
