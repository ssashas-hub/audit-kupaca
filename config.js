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
  visitRadius: 200,   // metara: koliko daleko od objekta sme da se započne poseta
  maxAccuracy: 150    // metara: lošija GPS tačnost se ne prihvata

};
