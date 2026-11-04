// Service Worker para Mi FOP - Fibra Óptica del Perú
const CACHE_NAME = 'fop-cache-v1';

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

// Manejo de eventos de clic en notificaciones en el celular
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
      for (const client of clientList) {
        if (client.url && 'focus' in client) {
          return client.focus();
        }
      }
      if (clients.openWindow) {
        return clients.openWindow('./');
      }
    })
  );
});

// Evento Push para notificaciones remotas en segundo plano (con la app cerrada)
self.addEventListener('push', (event) => {
  let data = {
    title: '🔔 ¡No olvides pagar tu plan de internet!',
    body: 'Tienes un recibo pendiente en Fibra Óptica del Perú. Evita cortes y paga al instante con Plin o Yape.',
    icon: 'icon-192.png',
    badge: 'icon-192.png'
  };

  if (event.data) {
    try {
      const payload = event.data.json();
      data = { ...data, ...payload };
    } catch(e) {
      data.body = event.data.text() || data.body;
    }
  }

  event.waitUntil(
    self.registration.showNotification(data.title, {
      body: data.body,
      icon: data.icon,
      badge: data.badge,
      vibrate: [200, 100, 200],
      tag: 'recordatorio-deuda-fop',
      renotify: true
    })
  );
});

// Recordatorio periódico en segundo plano (Periodic Background Sync)
self.addEventListener('periodicsync', (event) => {
  if (event.tag === 'check-deuda-sync') {
    event.waitUntil(
      self.registration.showNotification('🔔 Recordatorio de Pago FOP', {
        body: 'Revisa tu estado de cuenta en Mi FOP. Evita el corte de tu servicio de fibra óptica y FiberTV.',
        icon: 'icon-192.png',
        badge: 'icon-192.png',
        vibrate: [200, 100, 200],
        tag: 'recordatorio-deuda-fop'
      })
    );
  }
});
