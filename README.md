# Audit kupaca

Aplikacija za komercijalistu na terenu. Radi na telefonu (uspravno), može da se doda na početni ekran, radi i bez mreže (izmene i fotografije čekaju u telefonu dok se veza ne vrati).

- Podaci: SharePoint lista „Lične karte kupaca” na sajtu Komercijala, preko Microsoft Graph-a.
- Prijava: službeni Microsoft nalog (@ekoline.rs). Vidi samo ono čemu korisnik već ima pristup na sajtu.
- Fotografije: biblioteka dokumenata sajta Komercijala, fascikla „Fotografije kupaca”, naziv `šifra-datumvreme-broj.jpg`.
- Bez `clientId` u `config.js` aplikacija radi u probnom režimu sa izmišljenim kupcima.

## Fajlovi

`index.html`, `app.css`, `app.js`, `config.js`, `sw.js`, `manifest.webmanifest`, `icon.svg`. Nema servera ni baze: samo statični fajlovi. `test.html` i `shot.html` služe samo za proveru, ne postavljaju se.

## Povezivanje (jednom)

1. **Registracija u Entra ID** (entra.microsoft.com → Applications → App registrations → New registration)
   - Name: `Audit kupaca`
   - Supported account types: samo ova organizacija (single tenant)
   - Redirect URI: platforma **Single-page application (SPA)**, adresa na kojoj će aplikacija stajati, npr. `https://<korisnik>.github.io/audit-kupaca/index.html`
   - Posle pravljenja: API permissions → Add → Microsoft Graph → Delegated → `Sites.ReadWrite.All` (i `User.Read`, već postoji) → **Grant admin consent**.
   - Kopirati **Application (client) ID** u `config.js` → `clientId`.
2. **Postavljanje fajlova** na statični hosting sa HTTPS-om (GitHub Pages za sada; kasnije VPS). HTTPS je obavezan zbog prijave, lokacije i kamere.
3. **Fascikla za fotografije**: u biblioteci Documents sajta Komercijala napraviti fasciklu `Fotografije kupaca`.

## Selidba na VPS

Prebaciti iste fajlove, dodati novu adresu kao Redirect URI u Entra ID. Ništa drugo se ne menja.
