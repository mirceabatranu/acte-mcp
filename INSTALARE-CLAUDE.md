# Instalarea acte-mcp, făcută de Claude Code

Acest fișier e pentru **agentul Claude Code** care instalează acte-mcp pe un calculator
din birou. Agentul face tot, **mai puțin parolele de aplicație Gmail**: pe acelea le
introduce utilizatorul, într-un terminal deschis pentru asta.

Utilizatorul pornește instalarea cu un mesaj de felul:

> Instalează acte-mcp după https://github.com/mirceabatranu/acte-mcp/blob/main/INSTALARE-CLAUDE.md
> Conturi Gmail: ... · Alte adrese ale cabinetului: ... · Raportul la: ... · De la data: ...

---

## Reguli pentru agent

- Vorbește cu utilizatorul **în română**, simplu. La fiecare pas spune pe scurt ce faci.
- **Parolele nu trec niciodată prin tine.** Nu le ceri în chat, nu le tastezi și nu le
  citești. Nu folosi variabile de mediu cu parole și nu scrie parole în fișiere. Dacă
  utilizatorul lipește o parolă în chat, spune-i să o revoce din contul Google și să
  genereze alta.
- Valorile (conturi Gmail, adresele cabinetului, adresa pentru raport, data de început)
  vin din mesajul utilizatorului. Dacă lipsește vreuna, **întreabă**; nu ghici.
- Nu modifica și nu șterge nimic pe serverul Z:\ în afară de ce face programul. Nu umbla
  în Gmail, nu opri flow-urile vechi, nu urca nimic pe GitHub.
- **Unde rulezi comenzile.** În aplicația desktop Claude (fila Code), shell-ul tău poate fi
  izolat: o instalare făcută acolo poate ajunge într-un AppData virtual pe care
  utilizatorul nu-l vede. Rulează **instalările** (pașii 2 și 3) în terminalul vizibil al
  utilizatorului, cu `run_in_terminal`, apoi verifică rezultatul tot de acolo. În Claude
  Code din terminal (CLI), shell-ul e cel real și poți rula direct.
- Toate comenzile sunt pentru **PowerShell**.

---

## Pasul 1 — Verificări

1. Windows: `[Environment]::OSVersion.VersionString`.
2. Python 3.10 sau mai nou: `python --version` (sau `py -3 --version`).
   - Dacă lipsește sau e mai vechi, **cere acordul**, apoi:
     `winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements`.
     După instalare, terminalul trebuie redeschis.
3. Serverul:
   - `Test-Path Z:\` și primele foldere: `Get-ChildItem Z:\ -Directory | Select-Object -First 15 Name`.
     Folderele clienților arată ca „12. Nume firmă”.
   - Calea de rețea a lui Z: `(Get-PSDrive Z).DisplayRoot`. Dacă iese o cale `\\server\share`,
     **folosește-o ca rădăcină** în loc de `Z:\`: merge și când unitatea Z: nu s-a reconectat.
     Verifică întâi cu `Test-Path` că se deschide.
   - Dacă Z: nu există, oprește-te și întreabă utilizatorul unde e serverul.
4. Excelul de clienți `Clienti_Contabilitate.xlsx`. Caută-l, în ordinea asta:
   - `Get-ChildItem "G:\My Drive" -Recurse -Depth 4 -Filter "Clienti_Contabilitate*.xlsx" -ErrorAction SilentlyContinue`
   - același lucru pe rădăcina serverului, cu `-Depth 2`.
   Dacă găsești mai multe, întreabă care e cel bun. Dacă nu găsești niciunul, întreabă;
   instalarea poate continua și fără, dar atunci toate actele ajung în „_De verificat”.
   Dacă fișierul e în Google Drive „doar online”, spune-i utilizatorului să-l facă
   „Disponibil offline” (clic dreapta în Explorer).

## Pasul 2 — Instalarea programului (în terminalul utilizatorului)

```powershell
python -m pip install --user uv
python -m uv tool install --force git+https://github.com/mirceabatranu/acte-mcp
python -m uv tool update-shell
```

Verifică: `Test-Path "$env:USERPROFILE\.local\bin\acte-mcp.exe"` și
`Test-Path "$env:USERPROFILE\.local\bin\acte-sync.exe"`. Ambele trebuie să fie `True`.
De aici încolo folosește calea completă `& "$env:USERPROFILE\.local\bin\acte-mcp.exe"`, ca să
nu depinzi de PATH-ul încă nereîncărcat. Mai jos e scris scurt `acte-mcp`.

## Pasul 3 — Configurarea (fără întrebări)

```powershell
acte-mcp configurare --radacina "<radacina>" --conturi "<cont1>,<cont2>" `
  --adrese-proprii "<adresa1>,<adresa2>" --excel "<cale Excel>" `
  --de-la "<AAAA-LL-ZZ>" --raport-catre "<adresa raport>" --mod-test da
```

- `--mod-test da`: cât timp merg și flow-urile vechi, actele aduse de acte-mcp se scriu
  separat, în `...\Acte contabile\AAAA\LL\_test mail` și în `_De verificat (test mail)`.
  Așa nu se amestecă cu cele aduse de flow-ul vechi și se pot compara.
- Arată utilizatorului configurarea afișată și cere-i confirmarea.

## Pasul 4 — Parolele de aplicație (le face UTILIZATORUL)

Explică-i utilizatorului ce are de făcut **pe fiecare cont Gmail** din configurare:

1. Gmail → ⚙ → *Vezi toate setările* → *Redirecționare și POP/IMAP* → **Activează IMAP** → Salvează.
2. https://myaccount.google.com/security → **Verificare în doi pași** trebuie să fie activată.
3. https://myaccount.google.com/apppasswords → nume „acte-mcp” → **Creează**. Copiază cele 16 litere.

Apoi deschide-i **un terminal separat**, în care doar el tastează:

- în aplicația desktop: `open_terminal_tab`, cu titlul „Parole acte-mcp”, apoi
  `run_in_terminal` în acel tab cu comanda de mai jos, pentru primul cont;
- în CLI: spune-i să deschidă o fereastră PowerShell nouă.

```powershell
& "$env:USERPROFILE\.local\bin\acte-mcp.exe" parola <cont>
```

Comanda așteaptă parola. Utilizatorul o lipește (nu se vede pe ecran) și apasă Enter.
Repetă pentru fiecare cont. **Nu citi ce scrie utilizatorul și nu-i cere să-ți spună
parola.** Când îți spune că a terminat, verifică doar dacă parolele există:

```powershell
acte-mcp parole-lipsa
```

Codul 0 înseamnă că toate sunt salvate. Pentru conturile marcate „LIPSESTE”, reia pasul 4.

## Pasul 5 — Testul

```powershell
acte-mcp test
```

Trebuie să se termine cu „Totul e in regula.”. Erori frecvente:

| Mesaj | Ce faci |
|---|---|
| `AUTHENTICATIONFAILED` / `Invalid credentials` / `Application-specific password required` | IMAP nu e activat sau parola e greșită. Utilizatorul verifică pașii 4.1–4.3 și salvează din nou parola |
| `Nu gasesc dosarul 'Toate mesajele'` | Gmail → Setări → *Etichete* → la „Toate mesajele” bifează **Afișează în IMAP** |
| `Serverul nu e accesibil` | Rădăcina greșită sau serverul nu e conectat; reia pasul 1.3 |
| Eroare la Excel | Calea greșită sau fișierul e doar online în Drive; reia pasul 1.4 |
| `Din Excel fara folder pe server` | Doar informare: firmele acelea nu au folder pe Z:\ sau numele diferă. Spune-le utilizatorului |

## Pasul 6 — Prima descărcare

```powershell
acte-mcp descarca
```

Prima rulare ia tot de la data de început și poate dura mult. Rulează comanda în terminal
și așteaptă să termine; nu o întrerupe. La final arată utilizatorului:

- rezultatul (câte mesaje noi, câte fișiere salvate, câte au mers la „De verificat”, câte erori);
- `acte-mcp raport --zi <data de azi>`: lista fișierelor pe clienți, ce e de verificat și
  linkurile de descărcat manual.

Erorile pe un singur mesaj nu opresc restul și se reiau automat la următoarea rulare.

## Pasul 7 — Sarcina automată

```powershell
acte-mcp instaleaza-sarcina --minute 5
schtasks /Query /TN "BBA Descarcare mail" /FO LIST
```

Pornește-o o dată de mână și verifică jurnalul:

```powershell
schtasks /Run /TN "BBA Descarcare mail"
Get-Content "$env:USERPROFILE\.acte-mcp\descarcare.log" -Tail 5
```

În jurnal trebuie să apară o linie nouă „Rulare: ... "rezultat": "ok"”. Așteaptă câteva
secunde înainte de citire. Spune-i utilizatorului că sarcina rulează cât timp el e
conectat în Windows. Dacă închide calculatorul, la pornire programul ia tot ce a venit
între timp.

## Pasul 8 — Legătura cu Claude (MCP)

```powershell
claude mcp add --scope user acte -- "$env:USERPROFILE\.local\bin\acte-mcp.exe"
claude mcp list
```

Dacă `claude` nu e în PATH (în aplicația desktop), spune-i utilizatorului să adauge
serverul din setările aplicației, cu aceeași comandă și calea de mai sus. Uneltele
`acte_*` apar după repornirea sesiunii.

## Pasul 9 — Rezumatul pentru utilizator

Spune-i, pe scurt:

- ce s-a instalat și unde (`%USERPROFILE%\.acte-mcp\` pentru configurare, jurnal și stare);
- că descărcarea rulează singură la 5 minute, **fără Claude**;
- că în fiecare dimineață, după ora 7, vine raportul la adresa aleasă, cu numărătoarea de
  control („mesaje în Gmail / procesate / neprocesate”). **Dacă într-o zi nu vine,
  descărcarea nu rulează**;
- că la orice eroare vine imediat un mail de alertă;
- că actele sunt deocamdată în subfolderele „_test mail”. După câteva zile de rapoarte bune:
  `acte-mcp configurare --mod-test nu`, apoi se oprește partea de Gmail din flow-urile vechi.
