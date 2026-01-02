# ⛵ Yacht Operations Data Project

## 📖 Descrizione
Questo progetto mostra come ho migliorato la gestione dei dati operativi di una flotta di yacht, sostituendo report cartacei inefficaci con un sistema digitale automatizzato.  
L'obiettivo principale è ridurre errori manuali, velocizzare le operazioni di inserimento dati e fornire dashboard facilmente interpretabili per decisioni rapide.

## ⚠️ Contesto e Problema
Nella sede operativa i report erano cartacei e non centralizzati, con inefficienze legate a errori di trascrizione e tempi di elaborazione lunghi.  
Ogni yacht generava dati separati, e non esisteva una visione integrata delle operazioni e dei profitti.

## 💡 Soluzione
- **Excel avanzato**: file centrali (REPORT) e template per singoli yacht (REPORT_OPERAZIONI) con automazioni VBA, convalida dati e formattazione condizionale per ridurre errori e ottimizzare inserimenti.  
- **Database centralizzato**: i dati dei vari REPORT vengono uniti tramite PowerQuery, aggiornando automaticamente le dashboard.  
- **Interfaccia user-friendly**: nei vari REPORT_OPERAZIONI sono presenti pulsanti per aggiungere, eliminare e inviare dati al file centrale.  
- **AppSheet**: applicativo mobile per inserimento dati anche fuori ufficio.  
- **PowerBI**: dashboard per analisi temporali e finanziarie, con grafici, KPI e filtri per facilitare la lettura dei dati e l’individuazione di trend.

## 🗂️ Struttura della repository
Data-Analysis-Project/<br>
│<br>
├─ Excel/<br>
│ ├─ REPORT.xlsx<br>
│ ├─ REPORT_OPERAZIONI.xlsx<br>
| ├─ popolamento_file_report.xlxs<br>
│ └─ README.md<br>
│<br>
├─ PowerBI/<br>
│ ├─ REPORT.pbix<br>
│<br>
├─ AppSheet/<br>
│ └─ README.md<br>
│<br>
├─ Presentazione/<br>
│ └─ Demo_PowerPoint.pptx<br>
│<br>
└─ README.md<br>

## ⚙️ Strumenti e Tecnologie
- Microsoft Excel (macro VBA, convalida dati, formattazioni condizionali)  
- PowerQuery per integrazione dati  
- PowerBI per dashboard interattive  
- AppSheet per raccolta dati mobile  
- PowerPoint per presentazione finale e storytelling

## 🎯 Valore del progetto
Questo progetto dimostra la capacità di trasformare processi manuali in workflow digitali affidabili, migliorando efficienza operativa, accuratezza dei dati e capacità decisionale basata su KPI chiari.

## 🔍 Dettagli Tecnici

### 📝 Excel
- Macro VBA per:
  - Creazione automatica di checkbox all’inserimento dello yacht
  - Cancellazione automatica di righe alla rimozione del nome dello yacht
  - Pulsanti “Aggiungi”, “Elimina riga” e “Invia al Sistema” in REPORT_OPERAZIONI
- Convalida dati su tipologia di yacht, porti e regioni
- Formattazioni condizionali sulle date e sullo stato dei dati
- Integrazione tramite PowerQuery tra REPORT locali e file centrale

### 📊 PowerBI
- Misure DAX per:
  - Fatturato cumulato per yacht/prodotto/periodo
  - Percentuali di variazione mese-su-mese e anno-su-anno
- Grafici principali:
  - Linee temporali (trend toccate e profitti)
  - Barre orizzontali (top porti e categorie)
  - Mappe geografiche
- KPI e filtri dinamici per mese, sede, porto

### 📱 AppSheet
- Interfaccia mobile per inserimento dati
- Sincronizzazione con Excel centrale

## 🗒️ Note
- Tutti i dati reali sono stati anonimizzati o generati in maniera fittizia.  
- Le macro in Excel sono commentate per spiegare logica e funzionalità.  
- Il progetto può essere esteso o adattato facilmente ad altre realtà operative.
