# Frysta Word-diff-fall (schema 1)

`wordlevel_frozen_v1.json` innehåller **syntetiska, deterministiska** förslag
med fast gammal text, ny text och runs. De är härledda från de felklasser vi
undersöker, men är **inte** arkiverade AI-svar från rapportkörningen.

Varje förslag ger ett nytt DOCX-testdokument. Förslaget går genom verklig
`ChangePlanner`, `TrackedChangesRenderer` och `CommentsRenderer`. Verifiering:

- exakt gammal text när revisioner avvisas och ny text när de godtas,
- rätt fotnotsreferenser och bevarade positioner vid avvisning,
- en kommentar med alla tre förankringsmarkörer,
- inga fotnotsreferenser inuti revisionsomslag och inga dubblett-ID:n,
- fallback-orsak, antal revisionsomslag och antal markerade tecken.

Jämför aktuell arbetskatalog med en **redan committad basversion**:

```powershell
& "C:\temp\langimprover_venv\Scripts\python.exe" scripts/frozen_wordlevel_compare.py --baseline-ref HEAD --output frozen_wordlevel_comparison.json
```

Kommandot använder `git archive HEAD` till en temporär katalog för basversionens
kod; samma frysta fall och testverktyg används på båda sidor. Originaldokument
modifieras inte. Om en ny Git-commit görs först, välj dess föräldraref
(`HEAD~1`) eller en tidigare fast commit-hash för en meningsfull jämförelse.

**Tolkning:** `delta_revised_chars < 0` betyder att färre tecken omsluts av
`w:ins` och `w:del` i kandidatversionen. Det är ett precisionsmått, inte i sig
ett fullständigt visuellt kvalitetsbetyg. En förbättring i en syntetisk fixture
är inte automatiskt en förbättring i alla verkliga AI-förslag. Nästa steg är att
fylla på med exakt frysta, verkliga förslag från en och samma AI-körning.

## Verkliga textpar och opt-in-fångst (version 15)

Fixturefilen innehåller nu även **sju verkliga textpar** hämtade från
`Testrapport_2025_3.docx` och den godtagna texten i Word-körningen
`test_wordlevel_version12.docx`. Originalets run-gränser eller hela
styckets OOXML har behållits. Dessa är **inte** identifierade som de nio
verkliga fallen med `mixed_run_formatting` eller `multiple_edit_segments`;
Word-körningens statistik identifierar inte enskilda ursprungliga förslag.

För att frysa *just de verkliga förslag* som orsakar dessa två fallback-typer,
aktivera exporten i samma PowerShell-session före en vanlig Word-körning:

```powershell
$env:JBG_WORDLEVEL_CAPTURE_PATH = Join-Path (Get-Location) "wordlevel_capture_local.json"
# Kör därefter Word-granskningen precis som vanligt, med ordnivå aktivt.
Remove-Item Env:JBG_WORDLEVEL_CAPTURE_PATH -ErrorAction SilentlyContinue
```

Om några entydiga, oförändrade stycken kunde fångas skapas JSON-filen.
Den innehåller **hela originalstyckets text och XML samt AI-förslaget**.
Behandla filen som konfidentiell och committa den inte. Befintlig fil skrivs
aldrig över; ta bort/flytta den medvetet före en ny fångstkörning.

Jämför en sådan fångst mot en committad basversion med exakt samma indata:

```powershell
& "C:\temp\langimprover_venv\Scripts\python.exe" scripts/frozen_wordlevel_compare.py --baseline-ref HEAD --fixture-file wordlevel_capture_local.json --output frozen_wordlevel_real_local.json
```

Samma OOXML och förslag körs i båda versionerna. Notera att exporten endast
omfattar `paragraph` i huvudtexten med unik, exakt textankare och utan tidigare
revisioner i stycket. Därför kan **färre** fall fångas än rapportens totala
fallback-antal. Att ett fall saknas är inte ett tecken på att det är löst.
