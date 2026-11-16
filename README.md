# Convertiasc

Convertitore di file gratuito che funziona **interamente nel browser**: nessun file viene caricato su un server, nessuna installazione.

Include anche **Rimuovi bianco**, per rendere trasparente lo sfondo bianco (o di un altro colore) delle immagini.

## Cosa converte

| Tipo | Da | A |
|---|---|---|
| Immagini | PNG, JPG, WEBP, BMP, GIF, ICO, SVG, AVIF, HEIC\*, TIFF\* | PNG, JPG, WEBP, BMP, ICO (multi-dimensione), PDF |
| PDF | PDF | PNG / JPG (una per pagina), TXT, Markdown, HTML |
| Documenti | DOCX, Markdown, HTML, TXT | PDF, DOC (Word), HTML, Markdown, TXT |
| Dati e fogli di calcolo | CSV, TSV, XLSX, XLS, ODS, JSON, XML, YAML | XLSX, CSV, TSV, ODS, JSON, XML, YAML, HTML, Markdown |
| Audio | MP3, WAV, OGG, M4A, AAC, FLAC, OPUS | MP3, WAV, WEBM/OGG |
| Video | MP4, WEBM, MOV, MKV | MP3 / WAV (solo audio), WEBM, PNG / JPG (fotogramma) |

\* HEIC e TIFF si aprono solo nei browser che li supportano (es. Safari).

### Opzioni
- qualità JPG/WEBP e larghezza massima delle immagini
- unione di più immagini in un unico PDF
- bitrate MP3
- rimozione del bianco applicata direttamente durante la conversione
- download di tutti i risultati in un unico ZIP

## Rimuovi bianco
- **Tolleranza**: quanto un colore può essere lontano dal bianco ed essere comunque rimosso
- **Morbidezza bordi**: sfuma i bordi invece di tagliarli netti
- **Solo lo sfondo / tutto il bianco**: con la prima opzione viene rimosso solo il bianco collegato ai bordi dell'immagine, così il bianco all'interno (occhi, scritte…) resta
- **Pulizia aloni**: elimina il bordino chiaro intorno ai soggetti
- clic sull'immagine per rimuovere un colore diverso dal bianco
- anteprima su scacchiera, sfondo scuro o chiaro; tieni premuto per confrontare con l'originale
- download in PNG trasparente, singolo o tutte insieme in ZIP

## Usarlo

Apri `index.html` con un doppio clic. Al primo utilizzo di alcune funzioni (PDF, Excel, MP3…) serve internet per scaricare le librerie necessarie.

### Pubblicarlo con GitHub Pages
1. Crea un repository chiamato `Convertiasc` e caricaci questi file.
2. Vai su **Settings → Pages**.
3. In *Source* scegli **Deploy from a branch**, branch `main`, cartella `/ (root)`, e salva.
4. Dopo un minuto il sito sarà su `https://<tuo-utente>.github.io/Convertiasc/`.

## Struttura
```
index.html     pagina e interfaccia
css/style.css  stile (tema chiaro e scuro automatico)
js/app.js      tutta la logica di conversione
```

## Limiti
- Le conversioni video in WEBM e audio in WEBM/OGG avvengono in **tempo reale** (un video di 2 minuti richiede circa 2 minuti) e richiedono Chrome o Firefox.
- La conversione in PDF dei documenti mantiene titoli, paragrafi, elenchi e tabelle ma non immagini e impaginazione complessa.
- I file molto grandi (centinaia di MB) dipendono dalla memoria del browser.

## Librerie usate
[SheetJS](https://sheetjs.com), [jsPDF](https://github.com/parallax/jsPDF), [PDF.js](https://mozilla.github.io/pdf.js/), [Mammoth](https://github.com/mwilliamson/mammoth.js), [marked](https://marked.js.org), [Turndown](https://github.com/mixmark-io/turndown), [js-yaml](https://github.com/nodeca/js-yaml), [lamejs](https://github.com/zhuker/lamejs), [JSZip](https://stuk.github.io/jszip/).
