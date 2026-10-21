# Eredmény beíró – Android automatikus Google Táblázat-küldés

Ez a változat csak az Androidos alkalmazást módosítja. A képes eredmény mentése, az Excel-letöltés és a Messenger/Viber megosztás meglévő funkciói változatlanok.

## Mit csinál?
- A verseny utolsó futamának befejezésekor automatikusan sorba állítja a versenyeredményeket.
- A Google Táblázatban külön `Versenyeredmények` és `Összesített ranglista` munkalapot tart fenn.
- A ranglista versenyzőnként a korábbi eredmények közül a legjobb végeredményt tartja meg.
- Internetkimaradáskor a küldés helyben várakozik, és a program újrapróbálkozik, amikor a kapcsolat visszatér.

## Egyszeri Google-beállítás
1. Hozz létre egy Google Táblázatot.
2. A táblázatban nyisd meg: **Bővítmények → Apps Script**.
3. Nyisd meg a projektben lévő `GoogleAppsScript.gs` fájlt, másold a teljes tartalmát az Apps Script szerkesztőbe.
4. A kódban cseréld le a `SECRET` értékét egy saját, hosszú kulcsra. Példa helyett saját kulcsot használj.
5. Mentsd el a projektet.
6. Kattints **Telepítés → Új telepítés**. Típus: **Webalkalmazás**. Futtatás mint: **Én**. Hozzáférés: **Bárki** (különben az Android alkalmazás nem tudja elküldeni az eredményeket). Telepítsd, és engedélyezd a hozzáférést.
7. Másold ki a webalkalmazás `https://script.google.com/macros/s/.../exec` címét.
8. Az Android alkalmazás kezdőképernyőjén nyisd le a **Google Táblázat automatikus küldés beállítása** részt. Illeszd be az URL-t és pontosan ugyanazt a `SECRET` kulcsot, majd mentsd el.

A Google Apps Script nyilvános webalkalmazásának URL-je és kulcsa csak olyan személynek legyen átadva, akiben megbízol. A kulcs a kliensalkalmazásban megadva nem tekinthető teljesen titkosnak; elsősorban az illetéktelen, véletlen küldések ellen véd.

## Android APK
A GitHub Actions munkafolyamata a `build-apk.yml` vagy a `.github/workflows/main.yml` alapján készíti el az APK-t. Feltöltés után nyisd meg a GitHub tárhely **Actions** fülét, válaszd ki a sikeresen lefutott `APK építés` munkát, majd az **Artifacts** résznél töltsd le az `Eredmenytablazat-APK` csomagot.
