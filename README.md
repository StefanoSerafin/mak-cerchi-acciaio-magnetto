# Cerchi in acciaio Magnetto - consultazione per l'agente

Web app single-file per interrogare da iPhone e iPad il catalogo applicazioni dei
cerchi in acciaio Magnetto Wheels (`Catalogo Tecnico Magnetto Acciaio.xlsx`, "MW
Master Katalog" edizione maggio 2025) senza aprire l'Excel. Si sblocca con un PIN,
poi si filtra per vettura, misura del cerchio, pneumatico o codice.

Stesso schema di `../mak-distanziali-xilema/`: template + `data/build.py` +
`index.html` generato con il dataset cifrato.

**Online:** https://stefanoserafin.github.io/mak-cerchi-acciaio-magnetto/

## Accesso (PIN)

- Il `index.html` contiene solo il dataset cifrato (AES-GCM, chiave derivata dal PIN
  con PBKDF2). Con il PIN giusto i dati si decifrano nel browser.
- "Ricorda su questo dispositivo" salva il PIN in `localStorage`, solo su quel
  dispositivo.
- Il PIN non è nel repo.
- La decifratura richiede `https://` (GitHub Pages) o `http://localhost`.

## Come si usa

- **Cerca**: testo libero su marca, modello, versioni, codice `R1-…`, codice MWD,
  EAN, codice Alcar, omologazione, misura pneumatico (anche scritta `2055516`).
- **Filtri**: marca, modello, diametro, canale, fori x interasse, ET, foro centrale,
  pneumatico, finitura; in più "vetture ancora in produzione" e "novità". Ogni
  tendina elenca solo i valori ancora possibili con gli altri filtri, col conteggio.
- **Vetture / Cerchi**: la stessa selezione vista come elenco di applicazioni (una
  scheda per vettura) o di articoli (una scheda per codice, con le vetture
  compatibili).
- **Dettagli**: versioni, omologazione, codici, peso, pezzi per pallet, nota
  originale e le altre vetture su cui monta lo stesso cerchio.
- **Riepilogo**: testo pronto da copiare e mandare al cliente.
- **Sconto cliente %**: campo in alto, resta salvato sul dispositivo. A zero si vede
  solo il listino.

## Prezzi

Il prezzo mostrato è il lordo di listino del catalogo Magnetto (colonna "gross
price", in vigore dal 01.03.2025), per cerchio, IVA esclusa. **Non sono i prezzi
MAK A1/A2/A3** del file `LISTINI ACCIAIO(A1+A2+A3).xlsx` del sito agenti: se
servono quelli, vanno agganciati per codice articolo in `data/build.py`.

## Struttura

| File | Ruolo |
|---|---|
| `index.html` | L'app completa, dataset cifrato inline. **Generato**, non modificare a mano. |
| `index.template.html` | Markup, CSS e JS. Segnaposto `/*DATA_BLOB*/`, `/*BUILD_META*/`, `/*GENERATO*/`. **Questo si modifica a mano.** |
| `data/build.py` | Legge l'`.xlsx`, normalizza, cifra col PIN, rigenera `index.html`. |
| `data/Catalogo Tecnico Magnetto Acciaio.xlsx` | Fonte. **Non versionato.** |
| `data/cerchi.plain.json` | Dataset in chiaro, copia di servizio. **Non versionato.** |
| `assets/mak.png` | Logo MAK. |

## Rigenerare dopo un catalogo nuovo

Serve: `pip3 install openpyxl cryptography`

1. Sostituisci `data/Catalogo Tecnico Magnetto Acciaio.xlsx`.
2. `MAGNETTO_PIN='<pin>' python3 data/build.py`
3. Controlla il riepilogo stampato: "note rimaste in tedesco" e "canali non
   riconosciuti" devono essere vuoti. Se non lo sono, aggiungi le voci a `NOTE_IT`
   o correggi `norm_canale()`.

Per provarla in locale: `python3 -m http.server 8765` e apri
`http://localhost:8765/`.

## Regole di parsing

Ultimo build (2026-10-08): 972 applicazioni, 351 articoli, 48 marche (14 di veicoli
commerciali); 3.054 righe vuote e 1 riga senza codice (la 1026) scartate.

- **Marca**: il suffisso `Nfz` / `Nfz.` diventa "(commerciali)".
- **Canale**: spazi ripuliti, `6 1/2J` → `6 1/2 J`, `5,00 B` → `5 B`. Il profilo
  (`J`, `JJ`, `JK`, `B`) resta.
- **ET**: numerico; `HMA 121,5` (ruote gemellate MAN TGE / Crafter) resta testo.
- **Date**: `05.10` → `05/2010`; fine vuota = in produzione.
- **Pneumatici**: corretti `205//70R15` e `18565R15`; il resto com'è.
- **Finitura**: `schwarz` → Nero, `silber` → Argento.
- **Note**: si usa la colonna italiana; dove c'è solo il tedesco (127 righe) la
  traduzione è in `NOTE_IT` e `traduci_nota()`, **redazionale, non di Magnetto**.
  L'originale si vede nei Dettagli.

I dati non vengono corretti. Restano com'è nel catalogo, riga per riga:

- stesso articolo con pesi diversi: R1-2056 (7,5 / 10,4 kg), R1-2020 (7,05 / 10,3),
  R1-2002 (6,8 / 7,7), R1-1996, R1-1910;
- R1-2063 con due prezzi (75,6 / 78,3); R1-1910 con due EAN; R1-2131 con foro 78 e 78,1;
- sette articoli con due codici Alcar;
- riga 21 (Audi A3 GY): fine produzione 11/2018 precedente all'inizio 04/2020;
- R1-1612 (Clio III): pezzi per pallet a 0, mostrato come non indicato.

Nella vista "Cerchi" la scheda di un articolo riporta i dati della prima riga che
lo contiene.
