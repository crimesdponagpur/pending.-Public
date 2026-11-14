# lib: ॲपसोबत ठेवलेल्या बाहेरच्या लायब्ररी

या फाइली CDN (jsdelivr / cdnjs) वरून न घेता ॲपसोबतच दिल्या जातात, म्हणजे बाहेरचा कोणीही त्यात बदल करू शकत नाही.
प्रत्येक फाइल npm वरील अधिकृत पॅकेजमधून जशीच्या तशी घेतली आहे (npm ने sha512 integrity तपासून दिलेली). बदल करू नका.

| फाइल | npm पॅकेज | परवाना | sha256 |
|---|---|---|---|
| xlsx-js-style-1.2.0.bundle.js | xlsx-js-style@1.2.0 dist/xlsx.bundle.js | Apache-2.0 | 1c7abf2993ff2cd61e508f9268e9acda0098c9796f3925d2ba0d2579072653e2 |
| jszip-3.10.1.min.js | jszip@3.10.1 dist/jszip.min.js | MIT OR GPL-3.0-or-later | acc7e41455a80765b5fd9c7ee1b8078a6d160bbbca455aeae854de65c947d59e |
| hyperformula-2.7.1.full.min.js | hyperformula@2.7.1 dist/hyperformula.full.min.js | GPL-3.0-only | d45a0027c60de598c0512c8b3664fa5577e3ac16139699ee97e76e6ae453c5b4 |
| pdfjs-3.11.174/pdf.min.js | pdfjs-dist@3.11.174 build/pdf.min.js | Apache-2.0 | 5b5799e6f8c680663207ac5b42ee14eed2a406fa7af48f50c154f0c0b1566946 |
| pdfjs-3.11.174/pdf.worker.min.js | pdfjs-dist@3.11.174 build/pdf.worker.min.js | Apache-2.0 | feabdf309770ed24bba31a5467836cdc8cf639c705af27d52b585b041bb8527b |
| qrcode-1.0.0.min.js | qrcodejs@1.0.0 qrcode.min.js | MIT | c541ef06327885a8415bca8df6071e14189b4855336def4f36db54bde8484f36 |

परवान्यांच्या प्रती `licenses/` मध्ये. तपासणी: `sha256sum lib/*.js lib/pdfjs-3.11.174/*.js` (वरील यादीशी जुळायला हवे; `tools/check.js` हे आपोआप तपासते).
