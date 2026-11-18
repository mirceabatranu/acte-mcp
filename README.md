# acte-mcp

Două lucruri într-un singur program:

1. **Descărcarea automată** a tuturor atașamentelor din căsuțele Gmail pe serverul local
   (Z:\), direct în folderul clientului, la luna de lucru. Rulează singură, din Task
   Scheduler, la 5 minute. **Nu depinde de Claude, de conectori sau de Google Drive.**
2. **Un server MCP** prin care Claude vede actele de pe server: caută clientul, listează
   actele lunii, le citește (text sau imagine) și le marchează după importul în SAGA.

## Cum funcționează descărcarea

- Se conectează la Gmail prin **IMAP**, cu o *parolă de aplicație*. Nu are limită de
  16 MB și nu trece prin Drive.
- Deschide căsuța **doar în citire**: mailurile **rămân necitite**. Nu mută, nu
  etichetează și nu șterge nimic.
- Caută în „Toate mesajele”, deci prinde și mailurile arhivate. Ia mesajele primite și
  forward-urile, nu și mailurile trimise de cabinet către alții.
- Desface și mesajele atașate ca `.eml` (forward ca atașament), cu atașamentele din
  interiorul lor.
- **Firma** se stabilește după adresa expeditorului, din `Clienti_Contabilitate.xlsx`
  (coloanele F și K). La forward contează expeditorul original. Dacă un domeniu are mai
  multe firme, mailul trebuie să o numească. Ce nu e sigur ajunge în
  `_De verificat\AAAA\LL`.
- **Luna**: luna în care s-a primit mailul minus 1 (mail din noiembrie → `2026\10`).
- **Calea**: `Z:\<nr>. <Firma>\Acte contabile\AAAA\LL\<data mail>_<nume original>`.
- Imaginile de semnătură, `.ics`, `.vcf`, `winmail.dat` și newsletterele se sar.

### Cum știi că nu se pierde nimic

- Fiecare mesaj e trecut într-o bază locală (`~\.acte-mcp\stare.sqlite`) abia **după** ce
  toate fișierele lui sunt scrise pe disc. Dacă pică internetul sau serverul, mesajul se
  reia la rularea următoare.
- La fiecare rulare se reiau și ultimele 3 zile: nu rămân goluri, iar dublurile se sar
  (după conținut, nu doar după nume).
- Fișierele nu se suprascriu niciodată.
- **Raport zilnic** pe mail (implicit la office@bilanzaudit.ro, după ora 7) cu o
  **numărătoare de control**: „ieri: 37 de mesaje în Gmail, 37 procesate, 0 neprocesate”.
  Mai conține fișierele salvate pe clienți, ce e de verificat, linkurile Drive/WeTransfer
  de descărcat manual și erorile.
- **Alertă imediată** pe mail la orice eroare (cel mult o dată la 2 ore).
- **Dacă într-o zi nu vine raportul, descărcarea nu rulează.** Verificați calculatorul
  serverului.

Limită reală: fișierele trimise ca **link** (Drive, WeTransfer) nu sunt în mail. Apar în
raport și în `acte_linkuri`, dar se descarcă manual.

## Instalare (pe calculatorul care e mereu pornit și vede Z:\)

### 1. Gmail: IMAP și parolă de aplicație (pe fiecare cont)

1. Gmail → Setări → *Redirecționare și POP/IMAP* → **Activează IMAP**.
2. Contul Google → *Securitate* → **Verificare în doi pași** activată.
3. Contul Google → *Securitate* → **Parole pentru aplicații** → creați una, de exemplu
   „acte-mcp”. Copiați cele 16 litere.

### 2. Programul

```powershell
python -m pip install --user uv
python -m uv tool install git+https://github.com/mirceabatranu/acte-mcp
python -m uv tool update-shell
```

Închideți și redeschideți PowerShell, apoi:

```powershell
acte-mcp configurare
```

Configurarea întreabă:
- radăcina serverului (recomandat calea de rețea, de ex. `\\NAS\Public`, în loc de `Z:\`);
- conturile Gmail;
- calea către Excelul de clienți (de ex. `G:\My Drive\...\Clienti_Contabilitate.xlsx`);
- data de început;
- adresa pentru raport.

Restul setărilor se schimbă în `%USERPROFILE%\.acte-mcp\config.json`: adrese proprii,
expeditori ignorați și `mapare_manuala` (adresă sau domeniu → folder, pentru excepții).

```powershell
acte-mcp parola bilanzaudit@gmail.com
acte-mcp parola officebbaaudit@gmail.com
acte-mcp test
```

Parola nu se vede când o tastați. Se salvează în Windows Credential Manager, nu într-un
fișier. `test` verifică serverul, Excelul (câți clienți au adrese și ce firme nu au
folder), IMAP-ul și trimiterea de mail.

### 3. Prima descărcare și sarcina automată

```powershell
acte-mcp descarca
acte-mcp instaleaza-sarcina --minute 5
```

Prima rulare ia tot de la data de început și poate dura. Sarcina **„BBA Descarcare
mail”** rulează fără fereastră. Jurnalul este în `%USERPROFILE%\.acte-mcp\descarcare.log`.

```powershell
acte-mcp raport
```

arată raportul pe ieri fără să-l trimită (`--trimite` îl și trimite).

### 4. Claude (pe orice calculator care vede Z:\)

```powershell
claude mcp add --scope user acte -- "$env:USERPROFILE\.local\bin\acte-mcp.exe"
```

Pe calculatoarele care doar citesc actele nu e nevoie de parole și nici de sarcină. E
suficient un `config.json` cu `radacina` (și `excel_clienti`, pentru CUI și adrese).

## Uneltele MCP

| Unealtă | Ce face |
|---|---|
| `acte_status` | Serverul e accesibil? Când a rulat ultima descărcare, câte erori |
| `acte_clienti` | Folderele clienților, cu CUI și adrese de mail |
| `acte_lista` | Actele unui client pe o lună de lucru, cu originea (expeditor, subiect) și marcajul „importat_saga” |
| `acte_citeste` | Conținutul unui act: text (PDF, Excel, CSV, XML e-Factura, Word) sau imagine (poze, PDF scanat) |
| `acte_cauta` | Caută după numele fișierului, expeditor sau subiect |
| `acte_linkuri` | Mailurile cu linkuri Drive/WeTransfer de descărcat manual |
| `acte_marcheaza` | Marchează un act `importat_saga` / `verificat` / `ignorat` / `anulat` (se vede de pe toate calculatoarele) |
| `acte_muta` | Mută un act la alt client sau altă lună (de ex. din `_De verificat`), **în doi pași** |
| `acte_descarca_acum` | Pornește descărcarea imediat (doar pe calculatorul serverului) |

### Exemplu de flux cu SAGA

> „Pentru Alfa, ia actele din luna de lucru care nu sunt importate, citește facturile și
> extrasele, pregătește importul în SAGA, arată-mi ce urmează să import, iar după
> import marchează-le.”

Claude face: `acte_lista("Alfa", doar_neimportate=true)` → `acte_citeste(...)` pentru fiecare
act → `saga_pregateste_import` / `saga_pregateste_extras` / `saga_pregateste_bonuri` →
confirmarea dumneavoastră → `saga_importa` → `acte_marcheaza(cale, "importat_saga")`.

## Trecerea de la flow-ul vechi

Când rapoartele zilnice arată că descărcarea merge:
- opriți partea de Gmail din „Salvare acte office” / „Salvare acte Paul” și scriptul Apps
  Script care copia în Drive; altfel actele ajung de două ori pe Z:\ (o dată prin Drive);
- sarcina „BBA Copiere Drive - NAS” poate rămâne pentru ce se pune manual în Drive.

## Dezvoltare

```powershell
python -m venv .venv; .venv\Scripts\pip install -e . pytest
.venv\Scripts\python -m pytest -q
```

Testele simulează Gmail (IMAP) și serverul, fără internet.
