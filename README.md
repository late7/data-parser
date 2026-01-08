# Due Diligence Document Processor

Tämä työkalu lukee kaikki dokumentit `docs`-hakemistosta, analysoi ne OpenAI API:n avulla ja luo konsolidoidun Due Diligence -raportin JSON- ja DOCX-muodoissa.

## Ominaisuudet

- **Tukee useita tiedostomuotoja**: PDF, Excel (.xlsx), PowerPoint (.pptx)
- **AI-pohjainen analyysi**: Käyttää OpenAI GPT-4o -mallia tarkkaan tiedon poimintaan
- **Hallusinaatioiden minimointi**:
  - Matala temperature (0.1) API-kutsuissa
  - Eksplisiittiset ohjeet olla keksimättä tietoa
  - Jokaiseen tiedonpalaan lähdeviittaus
- **Rakenteellinen tuloste**:
  - JSON-tiedosto ohjelmalliseen käsittelyyn
  - DOCX-tiedosto ihmisen luettavaksi
- **Vain relevantti tieto**: Sisällyttää vain ne kohdat, joihin dokumenteista löytyy tietoa

## Vaatimukset

Asenna tarvittavat kirjastot:

```bash
pip install -r requirements.txt
```

## Konfiguraatio

Luo `.env`-tiedosto projektihakemistoon ja lisää OpenAI API-avain:

```
OPENAI_API_KEY=your-api-key-here
```

## Käyttö

1. Laita analysoitavat dokumentit `docs`-hakemistoon
2. Varmista että `master-document-template.json` on olemassa
3. Aja skripti:

```bash
python process_documents.py
```

## Tuloste

Skripti luo kaksi tiedostoa:

1. **consolidated_due_diligence.json** - Rakenteellinen JSON-dokumentti
   - Sisältää kaikki poimitut tiedot
   - Jokaisessa tiedossa lähdeviittaus
   - Confidence-taso merkitty

2. **consolidated_due_diligence.docx** - Ihmisluettava Word-dokumentti
   - Sama sisältö kuin JSON
   - Muotoiltu selkeästi otsikoilla ja luetteloilla
   - Lähdeviittaukset näkyvissä

## Rakenne

### master-document-template.json

Template määrittelee:
- **update_rule**: Miten tietoa käsitellään (append, overwrite, locked)
- **instruction**: Ohjeet AI:lle kyseisen kentän täyttämiseen

### Skriptin toiminta (Versio 3)

1. Lukee kaikki dokumentit `docs`-hakemistosta
2. Poimii tekstin jokaisesta dokumentista (PDF, Excel, PowerPoint)
3. Lähettää KOKO dokumentin OpenAI API:lle analysoitavaksi (yksi kutsu per dokumentti)
4. API palauttaa kaikki relevantit tiedot hierarkisessa muodossa
5. Yhdistää kaikkien dokumenttien tiedot säilyttäen template-järjestyksen
6. Tallentaa tulokset JSON- ja DOCX-muodoissa

**Parannettu versio 3:**
- Yksi API-kutsu per dokumentti (aiemmin 16 kutsua per dokumentti)
- Lähettää koko dokumentin (ei rajaa merkkimäärää)
- Käyttää korkeampaa max_tokens (8192) kattavampaan poimintaan
- Poistaa duplikaatit automaattisesti
- Säilyttää template-järjestyksen JSON-tulosteessa

## Turvallisuus

- API-avain tallennetaan `.env`-tiedostoon (ei versionhallintaan)
- Matala temperature API-kutsuissa minimoi hallusinaatiot
- Jokaisella tiedolla lähdeviittaus ja confidence-taso

## Huomioita

- API-kutsut maksavat (OpenAI hinnoittelu)
- Prosessointi vie aikaa (n. 1-2min per dokumentti)
- Pitkät dokumentit rajataan 30,000 merkkiin
- PDF-tiedostojen laatu vaikuttaa tekstin poimintaan

## Vianmääritys

**PDF-lukuvirheitä**:
- Tarkista että PDF ei ole suojattu
- Varmista että PDF sisältää tekstiä (ei pelkkiä kuvia)

**API-virheitä**:
- Tarkista API-avain `.env`-tiedostossa
- Varmista että sinulla on riittävästi OpenAI-krediittejä

**Unicode-virheitä konsolissa**:
- Windows-konsolin rajoitus
- Tiedostot luodaan silti onnistuneesti
