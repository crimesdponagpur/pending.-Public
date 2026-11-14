# प्रलंबित प्रकरण व्यवस्थापन प्रणाली

उपविभागीय कार्यालय नागपूर. ॲप: https://crimesdponagpur.github.io/pending.-Public/

यात फक्त ॲपचा code आहे. गुन्ह्यांची माहिती Google Sheet मध्ये, login मागे.

गोपनीय नोंदवही (फरार / पाहिजे आरोपी, निगराणी बदमाश, हद्दपार): https://crimesdponagpur.github.io/pending.-Public/nondvahi/ (माहिती फक्त स्वतःच्या Google Sheet मध्ये, login मागे)

ScanOS (कागद स्कॅन, PDF, Word, Excel, भाषांतर): https://crimesdponagpur.github.io/pending.-Public/scanner/ (कागद फोनमध्येच; OCR / भाषांतर स्वतःच्या Apps Script सर्व्हरवर)

## सुरक्षा (बदल करणाऱ्यांसाठी)

- `index.html` व `noting/index.html` मध्ये Content-Security-Policy आहे: फक्त ॲपचा स्वतःचा code चालतो, माहिती फक्त Apps Script सर्व्हरकडेच जाते.
- inline `<script>` मध्ये कोणताही बदल केल्यावर `node tools/check.js --fix` चालवा (CSP hash नवीन होतात), नाहीतर ॲप उघडणार नाही.
- `onclick="…"` सारखे inline event handler चालत नाहीत; JavaScript मधून `.onclick =` वापरा.
- बाहेरच्या लायब्ररी `lib/` मध्ये (CDN नाही). तपशील व sha256: `lib/README.md`.
- GitHub वर प्रत्येक बदलावर `MCR सुरक्षा तपासणी` आपोआप चालते (`.github/workflows/check.yml`).
- लॉगिन token व फोनमध्ये साठवलेली माहिती `vault.js` ने encrypt होते.
