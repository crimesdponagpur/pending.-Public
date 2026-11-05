# MCR Flutter डेमो

MCR प्रलंबित गुन्हे अॅपचा छोटा Flutter डेमो (नमुना माहिती). ३ स्क्रीन: मुख्य पान (कालावधीनुसार ४ टाइल), प्रलंबित यादी, गुन्हा तपशील (bottom sheet).

बिल्ड:
```
flutter create --platforms web .   # पहिल्यांदा, बाकी फाइल्स तयार करण्यासाठी
flutter build web --release --no-web-resources-cdn --pwa-strategy=none
```
Android साठी: `flutter create --platforms android .` नंतर `flutter build apk`.
