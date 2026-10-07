# Framework FUSION: Modellistica Stocastica e Predizione di Eventi ⚽📊
<img width="1920" height="1032" alt="{492ABFB2-3F38-49FE-8E67-F9FA75D92ED6}" src="https://github.com/user-attachments/assets/7e4023af-8b7d-473b-9c56-16a384ef5cc1" />

Un framework econometrico e software desktop integrato per la previsione di eventi discreti e il calcolo probabilistico nei sistemi competitivi bivariati. Il progetto si basa interamente sul mio impianto metodologico e di ricerca accademica: "Modelli probabilistici integrati per la previsione di eventi discreti: Poisson, regressione bayesiana, correzioni Dixon-Coles e indicatori dinamici in un framework FUSION".

---

## 🔒 Proprietà Intellettuale e Rilascio del Codice
**Nota sul codice sorgente:** Questo repository funge esclusivamente da **portfolio architetturale, documentale e scientifico**. Il codice sorgente completo in Visual Basic 6 (VB6) e l'eseguibile del software sono **privati** e tenuti riservati per proteggere l'algoritmo proprietario di calibrazione numerica e la proprietà intellettuale del software. 

Viene fornito di seguito l'impianto matematico-teorico completo per dimostrare l'ingegnerizzazione logica del sistema.

---

## 🔬 L'Architettura del Framework FUSION

Il modello supera i limiti della letteratura statistica classica strutturando il calcolo probabilistico su un'architettura a **livelli sequenziali**:

1. **Indicatori Strutturali (Forza Statica):** Calcolo della capacità di produzione/prevenzione tramite vettori di Forza d'Attacco (AS) e Forza di Difesa (DS), normalizzati rispetto alle medie del sistema, uniti all'Indice dei Punti per Gara (Ep) e alla Media Inglese (Fmi).
2. **Indicatori Dinamici (Forza Temporale):** Integrazione del rating Elo Relativo (aggiornato a ogni evento su coefficiente di sensibilità K), Media Mobile della Forma Recente (ultime N prove) ed Expected Points cumulativi (xP).
3. **Vantaggio Sistemico:** Modellazione del Fattore Campo Dinamico (Fcampo), non costante ma dipendente dallo squilibrio di punti e dalle differenze di rendimento casa/trasferta delle due entità.
4. **Shrinkage Bayesiano:** Stabilizzazione dei parametri di intensità grezzi (lambda_grezzo). Attraverso un Prior Informativo Gamma combinato con una Likelihood di Poisson, la stima viene ricondotta verso la media del sistema (mu) quando i dati storici sono scarsi, riducendo drasticamente la varianza nelle prime giornate.
5. **Poisson Bivariata Corretta:** Generazione della matrice delle probabilità dei punteggi integrando le correzioni moltiplicative locali di Dixon-Coles (rho = 0.05) per correggere la sovrastima dei pareggi bassi (0-0, 1-1) e inserire le dipendenze tattico-psicologiche.
6. **Correzioni di Struttura Aggiuntive:** Modulazione moltiplicativa per l'effetto di forte squilibrio tra intensità e per scenari a "Somma Alta" (quando la somma dei gol attesi è maggiore di 4) per favorire tabellini ad alto punteggio.
7. **Aggregazione Lineare e Logica Decisionale Multistrato:** Trasformazione della matrice congiunta troncata a supporto finito (matrice 7x7, da 0 a 6 gol) nelle probabilità aggregate di mercato (1-X-2, Over/Under, Goal/NoGoal) e loro classificazione tramite regole basate su soglie di confidenza e gap nativi.

---

## 📈 Formule Matematiche Chiave Implementate nel Software

### 1. Stabilizzazione Bayesiana dell'Intensità
Formula per il calcolo robusto di Lambda Finale (evita distorsioni nei dataset piccoli):
* **Modello:** Lambda_Finale = [ N / (N + N0) * Lambda_Ultra ] + [ N0 / (N + N0) * Media_Sistema ]
* *Significato:* Garantisce la stabilità del software anche nelle prime giornate di campionato grazie al meccanismo di shrinkage probabilistico.

### 2. Correzione Dixon-Coles per Dipendenze Locali
Modifica mirata per i punteggi più frequenti dove la Poisson classica fallisce:
* **Punteggi (0-0) e (1-1):** La probabilità viene ridotta applicando il fattore correttivo locale (Fattore = 1 - rho).
* **Punteggi (1-0) e (0-1):** La probabilità viene aumentata applicando il fattore correttivo locale (Fattore = 1 + rho).
* *Parametro empirico utilizzato:* rho = 0.05.

### 3. Indicatori Avanzati di Struttura della Previsione
- **Sorpresa (S):** Misura quanto la moda (il punteggio esatto atteso) è debole rispetto alla distribuzione complessiva delle probabilità.
  * *Formula:* S = 1 - Probabilità_Punteggio_Atteso
- **Coerenza (C):** Distanza matematica tra la cella dominante e il secondo punteggio più probabile nella matrice.
  * *Formula:* C = Probabilità_Punteggio_Atteso - Massima_Probabilità_Altri_Punteggi

---

## 🎮 Il Software Desktop (Ottimizzazione Numerica)

L'applicazione (sviluppata in VB6 con Windows API native) agisce come motore di calcolo operando due compiti principali:
- **Calibrazione Numerica Totale (Tasto F8):** Esegue un ciclo di ottimizzazione stocastica (Random Search fino a 2000 iterazioni) simulando i parametri macroscopici del framework sul database storico Risultati.txt, minimizzando la funzione di costo basata sulla Log-Loss reale del sistema.
- **UI Dinamica Interamente Custom:** Tabelle e controlli ListView vengono gestiti bypassando i limiti di refresh di Windows; l'intero layout è responsive e calcola l'auto-scaling delle proporzioni dei font e delle griglie geometriche in tempo reale durante il resize della finestra.

---
*Sviluppato da Massimiliano G.*
