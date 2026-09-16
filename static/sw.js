// Service worker mínimo: no cachea nada (TANA depende de conexión en vivo
// para procesar monografías), pero su sola presencia + registro es lo que
// el navegador exige para permitir "Agregar a pantalla de inicio".
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (event) => event.waitUntil(self.clients.claim()));
self.addEventListener("fetch", () => {
  // Passthrough: deja que cada request vaya normal a la red.
});
