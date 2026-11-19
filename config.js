// Podešavanja aplikacije. clientId se upisuje posle registracije u Entra ID.
// Ako je clientId prazan, aplikacija radi u probnom režimu sa izmišljenim kupcima.
window.APP_CONFIG = {
  tenantId: 'deadc661-499c-4753-9961-d01d78ebd823',
  clientId: '911a5d3b-8575-40d9-9307-5e07619e6ae7',
  spHost: 'ekolineks.sharepoint.com',
  sitePath: '/sites/Komercijala',
  listName: 'Lične karte kupaca',
  photoFolder: 'Fotografije kupaca',
  logList: 'Dnevnik rada',
  visitRadius: 50,    // metara: krug oko objekta (dodaje se trenutna GPS preciznost)
  maxAccuracy: 100,   // metara: lošija GPS preciznost se ne prihvata za početak posete
  allowNoGps: true,  // PROBA: kad GPS ne uspe, nudi „Nastavi bez GPS-a” (beleži se u Dnevnik rada). Za pravi rad staviti false.
  mapAccuracy: 10,    // metara: potrebna preciznost kada se objekat prvi put mapira (ispred ulaza)
  // uloge: admin (vidi i podešava sve), kom (komercijalista), op (call centar), mag (magacin)
  roles: {
    'sasa@ekolineks.onmicrosoft.com': 'admin',
    'powerapps@ekolineks.onmicrosoft.com': 'admin',
    'petarjovanovic@ekoline.rs': 'kom',
    'goca@ekolineks.onmicrosoft.com': 'op',
    'tamara@ekolineks.onmicrosoft.com': 'op',
    'magacin@ekoline.rs': 'mag'   // nema licencu: ne prijavljuje se, poruke mu idu na mejl
  },
  names: {
    'sasa@ekolineks.onmicrosoft.com': 'Aleksandar',
    'powerapps@ekolineks.onmicrosoft.com': 'Admin (server)',
    'petarjovanovic@ekoline.rs': 'Petar',
    'goca@ekolineks.onmicrosoft.com': 'Goca',
    'tamara@ekolineks.onmicrosoft.com': 'Tamara',
    'magacin@ekoline.rs': 'Magacin'
  },
  defaultRole: 'op',   // nalog koji nije na spisku vidi aplikaciju kao call centar (ne kao komercijalista)
  reportTo: 'aleksandarmilojevic@ekoline.rs'   // kome ide dnevni izveštaj na kraju dana
};
