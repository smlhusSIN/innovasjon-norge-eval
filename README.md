# Innovasjon Norge & NIC Søknadsevaluering

Dette verktøyet lar deg enkelt evaluere søknader til Innovasjon Norge og NIC (Norwegian Innovation Clusters) – rett fra nettleseren din. Du laster opp en PDF, velger type evaluering, og får en ferdig rapport i Excel på sekunder.

---

## Hvordan fungerer det?

1. **Du åpner nettsiden** (lokalt på din PC).
2. **Du laster opp en PDF med søknaden din.**
3. **Du velger hvilken type evaluering du vil ha:**
   - Oppstart 1, 2 eller 3 (Innovasjon Norge)
   - NIC Klyngeevaluering
4. **Trykk på knappen.**
5. **AI-en leser og vurderer søknaden automatisk** etter forhåndsdefinerte kriterier.
6. **Du får en Excel-rapport** med både poeng og kommentarer – klar til bruk!

---

## Slik setter du opp løsningen

1. **Klon prosjektet:**
   ```bash
   git clone <repo-url>
   cd innovasjon-norge-evaluering
   ```
2. **Installer det du trenger:**
   ```bash
   pip install -r requirements.txt
   ```
3. **Legg inn OpenAI-nøkkelen din:**
   - Lag en fil som heter `.env` i prosjektmappen.
   - Skriv inn:
     ```
     OPENAI_API_KEY=din-api-nøkkel-her
     ```

---

## Slik bruker du løsningen

1. **Start serveren:**
   ```bash
   uvicorn app:app --reload
   ```
2. **Gå til** [http://localhost:8000](http://localhost:8000) i nettleseren.
3. **Last opp PDF og velg evalueringstype.**
4. **Trykk på "Evaluer og last ned rapport".**
5. **Excel-rapporten lastes ned automatisk.**

---

## Hva skjer i bakgrunnen?

- PDF-en du laster opp blir lest og tekstinnholdet hentes ut.
- AI (OpenAI GPT-4o) vurderer søknaden etter relevante kriterier for valgt regime.
- Resultatene samles og det lages en Excel-rapport med både poeng, kommentarer og sammendrag.
- Rapporten sendes rett tilbake til deg – ingen data lagres permanent.

---

## Sikkerhet og personvern

- **API-nøkkelen** din er kun lagret lokalt i `.env`-filen.
- **PDF-filer** slettes etter bruk og lagres aldri permanent.
- **Ingen sensitive data** sendes til andre enn OpenAI (for selve vurderingen).

---

## Vil du endre eller utvide?

- **Nye evalueringsregimer:**
  - Legg til nye funksjoner i `evaluate_application.py` eller `evaluate_nic_application.py`.
  - Legg til valget i dropdownen i `app.py`.
- **Endre kriterier:**
  - Rediger spørsmålene i de samme filene.
- **Støtte for flere filtyper:**
  - Utvid funksjonen `read_application_text`.

---

## Feilsøking

- **API-feil:** Sjekk at `.env`-filen har riktig OpenAI-nøkkel.
- **Excel-fil kan ikke lagres:** Lukk filen i Excel før du prøver igjen.
- **Ingen tekst funnet i PDF:** Sjekk at PDF-en faktisk inneholder tekst (ikke bare bilder).

---

## Kontakt og bidrag

- Spørsmål eller forslag? Opprett en issue eller ta kontakt!
- Bidrag og pull requests er alltid velkomne.

---

**Lykke til med søknadene!** 🚀 