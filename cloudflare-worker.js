// ============================================================
//  CLOUDFLARE WORKER — PWA Shell + Push Notifications
//  Bindings requeridos:
//    - PUSH_SUBSCRIPTIONS (KV namespace)
//  Variables de entorno (secrets):
//    - VAPID_PUBLIC_KEY
//    - VAPID_PRIVATE_KEY
//    - PUSH_AUTH_TOKEN  (token secreto que GAS envía para autenticarse)
// ============================================================

const VAPID_PUBLIC  = 'BGivQjFutLF_ixAlil_Q2ntGtM1RgRcLEuxtlwXLknRXN_GOogO26oCOcm9aTfhYfrKPicrhUQP7AqBk4Q1PpRY';
const VAPID_SUBJECT = 'mailto:victorcaracela@gmail.com';
const GAS_URL       = 'https://script.google.com/macros/s/AKfycbwJrer0KO6jEd9HFso-AKzARyzlVdRrblJzm1H2i2ylWCbsCS9XzLGAfuQio2EPMzg/exec';

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // ---- CORS preflight ----
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders() });
    }

    // ---- API ROUTES ----
    if (url.pathname === '/api/push/subscribe' && request.method === 'POST') {
      return handleSubscribe(request, env);
    }
    if (url.pathname === '/api/push/send' && request.method === 'POST') {
      return handleSend(request, env);
    }
    if (url.pathname === '/api/push/send-bulk' && request.method === 'POST') {
      return handleSendBulk(request, env);
    }
    if (url.pathname === '/api/push/unsubscribe' && request.method === 'POST') {
      return handleUnsubscribe(request, env);
    }
    if (url.pathname === '/sw.js') {
      return serveServiceWorker();
    }
    if (url.pathname === '/api/push/vapid-key') {
      return json({ publicKey: VAPID_PUBLIC });
    }
    if (url.pathname === '/api/push/status' && request.method === 'POST') {
      return handleStatus(request, env);
    }

    // ---- Página principal (shell PWA) ----
    return serveShell();
  }
};

// ============================================================
//  SUBSCRIBE — Guardar suscripción push de un usuario
//  Body: { dni: "12345678", subscription: {...} }
// ============================================================
async function handleSubscribe(request, env) {
  try {
    const { dni, subscription } = await request.json();
    if (!dni || !subscription?.endpoint) {
      return json({ ok: false, error: 'dni y subscription requeridos' }, 400);
    }

    // Leer suscripciones existentes del usuario
    const key = `push:${dni}`;
    const existing = JSON.parse(await env.PUSH_SUBSCRIPTIONS.get(key) || '[]');

    // Evitar duplicados por endpoint
    const already = existing.find(s => s.endpoint === subscription.endpoint);
    if (!already) {
      existing.push(subscription);
      await env.PUSH_SUBSCRIPTIONS.put(key, JSON.stringify(existing));
    }

    return json({ ok: true, total: existing.length });
  } catch (e) {
    return json({ ok: false, error: e.message }, 500);
  }
}

// ============================================================
//  UNSUBSCRIBE — Eliminar suscripción de un usuario
//  Body: { dni: "12345678", endpoint: "https://..." }
// ============================================================
async function handleUnsubscribe(request, env) {
  try {
    const { dni, endpoint } = await request.json();
    if (!dni) return json({ ok: false, error: 'dni requerido' }, 400);

    const key = `push:${dni}`;
    const existing = JSON.parse(await env.PUSH_SUBSCRIPTIONS.get(key) || '[]');
    const filtered = existing.filter(s => s.endpoint !== endpoint);
    await env.PUSH_SUBSCRIPTIONS.put(key, JSON.stringify(filtered));

    return json({ ok: true });
  } catch (e) {
    return json({ ok: false, error: e.message }, 500);
  }
}

// ============================================================
//  SEND — Enviar push a UN usuario por DNI
//  Body: { token, dni, title, body, icon?, url?, tag? }
//  (llamado desde Google Apps Script)
// ============================================================
async function handleSend(request, env) {
  try {
    const data = await request.json();

    // Autenticar — GAS debe enviar token secreto
    if (!data.token || data.token !== env.PUSH_AUTH_TOKEN) {
      return json({ ok: false, error: 'No autorizado' }, 401);
    }

    const { dni, title, body, icon, url, tag } = data;
    if (!dni || !title) return json({ ok: false, error: 'dni y title requeridos' }, 400);

    const key = `push:${dni}`;
    const subs = JSON.parse(await env.PUSH_SUBSCRIPTIONS.get(key) || '[]');
    if (!subs.length) return json({ ok: true, sent: 0, reason: 'Sin suscripciones' });

    const payload = JSON.stringify({
      title,
      body: body || '',
      icon: icon || 'https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o',
      badge: 'https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o',
      url: url || '/',
      tag: tag || 'default'
    });

    const results = await sendToAll(subs, payload, env);

    // Limpiar suscripciones expiradas
    const valid = subs.filter((_, i) => results[i] !== 'gone');
    if (valid.length !== subs.length) {
      await env.PUSH_SUBSCRIPTIONS.put(key, JSON.stringify(valid));
    }

    const sent = results.filter(r => r === 'ok').length;
    return json({ ok: true, sent, total: subs.length });
  } catch (e) {
    return json({ ok: false, error: e.message }, 500);
  }
}

// ============================================================
//  SEND-BULK — Enviar push a VARIOS usuarios por DNI[]
//  Body: { token, dnis: ["123","456"], title, body, icon?, url?, tag? }
// ============================================================
async function handleSendBulk(request, env) {
  try {
    const data = await request.json();
    if (!data.token || data.token !== env.PUSH_AUTH_TOKEN) {
      return json({ ok: false, error: 'No autorizado' }, 401);
    }

    const { dnis, title, body, icon, url, tag } = data;
    if (!Array.isArray(dnis) || !title) return json({ ok: false, error: 'dnis[] y title requeridos' }, 400);

    const payload = JSON.stringify({
      title,
      body: body || '',
      icon: icon || 'https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o',
      badge: 'https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o',
      url: url || '/',
      tag: tag || 'default'
    });

    let totalSent = 0;
    for (const dni of dnis) {
      const key = `push:${dni}`;
      const subs = JSON.parse(await env.PUSH_SUBSCRIPTIONS.get(key) || '[]');
      if (!subs.length) continue;

      const results = await sendToAll(subs, payload, env);
      const valid = subs.filter((_, i) => results[i] !== 'gone');
      if (valid.length !== subs.length) {
        await env.PUSH_SUBSCRIPTIONS.put(key, JSON.stringify(valid));
      }
      totalSent += results.filter(r => r === 'ok').length;
    }

    return json({ ok: true, sent: totalSent, dnis: dnis.length });
  } catch (e) {
    return json({ ok: false, error: e.message }, 500);
  }
}

// ============================================================
//  STATUS — Verificar suscripciones de un usuario
//  Body: { token, dni }
// ============================================================
async function handleStatus(request, env) {
  try {
    const data = await request.json();
    if (!data.token || data.token !== env.PUSH_AUTH_TOKEN) {
      return json({ ok: false, error: 'No autorizado' }, 401);
    }
    const { dni } = data;
    if (!dni) return json({ ok: false, error: 'dni requerido' }, 400);

    const key = `push:${dni}`;
    const subs = JSON.parse(await env.PUSH_SUBSCRIPTIONS.get(key) || '[]');
    return json({ ok: true, dni, subscriptions: subs.length, endpoints: subs.map(s => s.endpoint?.substring(0, 60) + '...') });
  } catch (e) {
    return json({ ok: false, error: e.message }, 500);
  }
}

// ============================================================
//  WEB PUSH — Enviar notificación usando protocolo Web Push
// ============================================================
async function sendToAll(subscriptions, payload, env) {
  return Promise.all(subscriptions.map(sub => sendPush(sub, payload, env)));
}

async function sendPush(subscription, payload, env) {
  try {
    const vapidPrivate = env.VAPID_PRIVATE_KEY;
    const aud = new URL(subscription.endpoint).origin;

    // JWT Header + Claims
    const header = { typ: 'JWT', alg: 'ES256' };
    const now = Math.floor(Date.now() / 1000);
    const claims = { aud, exp: now + 86400, sub: VAPID_SUBJECT };

    // Firmar JWT con clave VAPID privada
    const jwt = await signJWT(header, claims, vapidPrivate);

    // Cifrar payload con claves del usuario
    const encrypted = await encryptPayload(subscription, payload);

    const resp = await fetch(subscription.endpoint, {
      method: 'POST',
      headers: {
        'Authorization': `vapid t=${jwt}, k=${VAPID_PUBLIC}`,
        'Content-Encoding': 'aes128gcm',
        'Content-Type': 'application/octet-stream',
        'TTL': '86400'
      },
      body: encrypted
    });

    if (resp.status === 201 || resp.status === 200) return 'ok';
    if (resp.status === 404 || resp.status === 410) return 'gone';
    return 'fail';
  } catch (e) {
    console.error('sendPush error:', e);
    return 'fail';
  }
}

// ============================================================
//  CRYPTO — JWT signing + Web Push Encryption
// ============================================================
function base64urlEncode(buf) {
  const bytes = buf instanceof ArrayBuffer ? new Uint8Array(buf) : buf;
  let str = '';
  for (const b of bytes) str += String.fromCharCode(b);
  return btoa(str).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function base64urlDecode(str) {
  str = str.replace(/-/g, '+').replace(/_/g, '/');
  while (str.length % 4) str += '=';
  const bin = atob(str);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

async function signJWT(header, claims, privateKeyBase64url) {
  const enc = new TextEncoder();
  const headerB64 = base64urlEncode(enc.encode(JSON.stringify(header)));
  const claimsB64 = base64urlEncode(enc.encode(JSON.stringify(claims)));
  const unsignedToken = `${headerB64}.${claimsB64}`;

  const keyData = base64urlDecode(privateKeyBase64url);

  // Importar clave privada ECDSA P-256
  const jwk = {
    kty: 'EC', crv: 'P-256',
    d: privateKeyBase64url,
    x: VAPID_PUBLIC ? base64urlEncode(base64urlDecode(VAPID_PUBLIC).slice(1, 33)) : '',
    y: VAPID_PUBLIC ? base64urlEncode(base64urlDecode(VAPID_PUBLIC).slice(33, 65)) : ''
  };

  const key = await crypto.subtle.importKey('jwk', jwk, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['sign']);
  const sig = await crypto.subtle.sign({ name: 'ECDSA', hash: 'SHA-256' }, key, enc.encode(unsignedToken));

  // Convertir DER signature a fixed-length r||s
  const sigBytes = new Uint8Array(sig);
  const r = sigBytes.slice(0, 32);
  const s = sigBytes.slice(32, 64);
  const fixedSig = new Uint8Array(64);
  fixedSig.set(r); fixedSig.set(s, 32);

  return `${unsignedToken}.${base64urlEncode(fixedSig)}`;
}

async function encryptPayload(subscription, payloadText) {
  const enc = new TextEncoder();
  const payload = enc.encode(payloadText);

  // Decodificar claves del cliente
  const clientPublicKey = base64urlDecode(subscription.keys.p256dh);
  const clientAuth = base64urlDecode(subscription.keys.auth);

  // Generar par de claves efímeras
  const localKeys = await crypto.subtle.generateKey({ name: 'ECDH', namedCurve: 'P-256' }, true, ['deriveBits']);
  const localPublicRaw = new Uint8Array(await crypto.subtle.exportKey('raw', localKeys.publicKey));

  // Importar clave pública del cliente
  const clientKey = await crypto.subtle.importKey('raw', clientPublicKey, { name: 'ECDH', namedCurve: 'P-256' }, false, []);

  // ECDH shared secret
  const sharedSecret = new Uint8Array(await crypto.subtle.deriveBits({ name: 'ECDH', public: clientKey }, localKeys.privateKey, 256));

  // Salt aleatorio
  const salt = crypto.getRandomValues(new Uint8Array(16));

  // HKDF para auth_info
  const authInfo = concat(enc.encode('WebPush: info\0'), clientPublicKey, localPublicRaw);
  const prk = await hkdfExtract(clientAuth, sharedSecret);
  const ikm = await hkdfExpand(prk, authInfo, 32);

  // HKDF para content encryption key
  const contentPrk = await hkdfExtract(salt, ikm);
  const cekInfo = enc.encode('Content-Encoding: aes128gcm\0');
  const cek = await hkdfExpand(contentPrk, cekInfo, 16);

  // HKDF para nonce
  const nonceInfo = enc.encode('Content-Encoding: nonce\0');
  const nonce = await hkdfExpand(contentPrk, nonceInfo, 12);

  // Padding + payload
  const paddedPayload = concat(payload, new Uint8Array([2])); // delimiter

  // AES-128-GCM encrypt
  const aesKey = await crypto.subtle.importKey('raw', cek, 'AES-GCM', false, ['encrypt']);
  const encrypted = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv: nonce }, aesKey, paddedPayload));

  // Header: salt(16) + rs(4) + idlen(1) + keyid(65) + encrypted
  const rsBytes = new Uint8Array(4);
  new DataView(rsBytes.buffer).setUint32(0, 4096); // standard record size
  const idlen = new Uint8Array([65]); // length of uncompressed P-256 public key

  return concat(salt, rsBytes, idlen, localPublicRaw, encrypted).buffer;
}

async function hkdfExtract(salt, ikm) {
  const key = await crypto.subtle.importKey('raw', salt, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  return new Uint8Array(await crypto.subtle.sign('HMAC', key, ikm));
}

async function hkdfExpand(prk, info, length) {
  const key = await crypto.subtle.importKey('raw', prk, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  const input = concat(info, new Uint8Array([1]));
  const result = new Uint8Array(await crypto.subtle.sign('HMAC', key, input));
  return result.slice(0, length);
}

function concat(...arrays) {
  const len = arrays.reduce((s, a) => s + a.length, 0);
  const result = new Uint8Array(len);
  let offset = 0;
  for (const a of arrays) { result.set(a, offset); offset += a.length; }
  return result;
}

// ============================================================
//  SERVICE WORKER (servido desde /sw.js)
// ============================================================
function serveServiceWorker() {
  const sw = `
self.addEventListener('push', function(event) {
  let data = { title: 'ADECCO ISOS', body: 'Nueva notificación', icon: '', badge: '', url: '/', tag: 'default' };
  try { data = Object.assign(data, event.data.json()); } catch(e) {}

  const options = {
    body: data.body,
    icon: data.icon || 'https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o',
    badge: data.badge || 'https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o',
    vibrate: [300, 100, 300, 100, 300, 100, 500],
    tag: data.tag || 'default',
    renotify: true,
    requireInteraction: true,
    silent: false,
    data: { url: data.url || '/' },
    actions: [
      { action: 'open', title: 'Abrir App' },
      { action: 'dismiss', title: 'Cerrar' }
    ]
  };

  event.waitUntil(self.registration.showNotification(data.title, options));
});

self.addEventListener('notificationclick', function(event) {
  event.notification.close();
  if (event.action === 'dismiss') return;
  const url = event.notification.data?.url || '/';
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function(clientList) {
      for (const client of clientList) {
        if (client.url.includes(self.location.origin) && 'focus' in client) {
          return client.focus();
        }
      }
      return clients.openWindow(url);
    })
  );
});
`;
  return new Response(sw, {
    headers: { 'Content-Type': 'application/javascript', 'Service-Worker-Allowed': '/', ...corsHeaders() }
  });
}

// ============================================================
//  SHELL PWA — Página principal con registro de push
// ============================================================
function serveShell() {
  const manifest = JSON.stringify({
    name: 'ADECCO - ISOS',
    short_name: 'ADECCO - ISOS',
    description: 'Sistema de Gestión Operacional SSOMA',
    start_url: '/',
    display: 'standalone',
    background_color: '#ffffff',
    theme_color: '#2c3e50',
    orientation: 'any',
    icons: [72,96,128,144,152,192,384,512].map(s => ({
      src: 'https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o',
      sizes: s+'x'+s, type: 'image/png', purpose: 'any'
    }))
  });

  const manifestB64 = btoa(unescape(encodeURIComponent(manifest)));

  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
  <title>ADECCO - ISOS</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@600;700&family=Orbitron:wght@700&display=swap" rel="stylesheet">
  <link rel="icon" type="image/png" href="https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o">
  <link rel="apple-touch-icon" sizes="180x180" href="https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o">
  <meta name="theme-color" content="#2c3e50">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
  <meta name="apple-mobile-web-app-title" content="ADECCO - ISOS">
  <meta name="mobile-web-app-capable" content="yes">
  <meta name="application-name" content="ADECCO - ISOS">
  <meta name="screen-orientation" content="any">
  <link rel="manifest" href="data:application/json;base64,${manifestB64}">
  <style>
    body,html{margin:0;padding:0;height:100vh;width:100vw;overflow:hidden;background:#fff}
    #splash{position:fixed;top:0;left:0;width:100%;height:100%;background:#fff;display:flex;flex-direction:column;justify-content:center;align-items:center;z-index:9999;transition:opacity .5s ease-out}
    #splash img{width:200px;height:auto;max-height:200px}
    #splash-text{font-family:'Orbitron','Poppins',sans-serif;font-size:22px;font-weight:700;margin-top:30px;text-align:center;letter-spacing:4px;background:linear-gradient(90deg,#8b0000,#f00,#8b0000);background-size:200% auto;-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;animation:gshift 3s ease infinite,pulse 2s ease-in-out infinite;text-transform:uppercase;position:relative}
    @keyframes gshift{0%{background-position:0% center}50%{background-position:100% center}100%{background-position:0% center}}
    @keyframes pulse{0%,100%{transform:scale(1);filter:brightness(1)}50%{transform:scale(1.02);filter:brightness(1.2)}}
    #splash-text::before{content:'';position:absolute;top:0;left:-100%;width:100%;height:100%;background:linear-gradient(90deg,transparent,rgba(255,255,255,.8),transparent);animation:shine 3s infinite}
    @keyframes shine{0%{left:-100%}50%,100%{left:100%}}
    .dots{display:inline-block;font-family:'Orbitron',monospace;font-weight:700}
    .dots::after{content:'';animation:dots 1.5s steps(4,end) infinite}
    @keyframes dots{0%{content:''}25%{content:'.'}50%{content:'..'}75%{content:'...'}100%{content:''}}
    iframe{width:100%;height:100%;border:none;opacity:0;transition:opacity .5s ease-in}
    iframe.loaded{opacity:1}
    @media(max-width:768px){#splash-text{font-size:18px;letter-spacing:3px}#splash img{width:160px;max-height:160px}}
  </style>
</head>
<body>
  <div id="splash">
    <img src="https://lh3.googleusercontent.com/d/15B-wj4iw5B7RpDXQg2mrdxTAn-3kfeMa" alt="Cargando">
    <div id="splash-text">CARGANDO SISTEMA<span class="dots"></span></div>
  </div>
  <!-- Banner de activación de notificaciones (requiere gesto del usuario) -->
  <div id="push-banner" style="display:none;position:fixed;bottom:20px;left:50%;transform:translateX(-50%);background:linear-gradient(135deg,#1a1a2e,#16213e);color:#fff;padding:14px 24px;border-radius:16px;z-index:10000;box-shadow:0 8px 32px rgba(0,0,0,0.3);align-items:center;gap:14px;font-family:'Poppins',sans-serif;font-size:14px;max-width:90vw;border:1px solid rgba(255,255,255,0.1)">
    <div style="background:#ffc107;color:#1a1a2e;width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0">
      <span>&#128276;</span>
    </div>
    <span style="flex:1">Activa las notificaciones para recibir alertas de EPP y capacitaciones</span>
    <button onclick="activarNotificaciones()" style="background:#ffc107;color:#1a1a2e;border:none;padding:8px 18px;border-radius:10px;font-weight:700;cursor:pointer;font-size:13px;white-space:nowrap">Activar</button>
    <button onclick="hideNotificationBanner()" style="background:transparent;color:#aaa;border:none;font-size:20px;cursor:pointer;padding:0 4px;line-height:1">&times;</button>
  </div>

  <iframe id="main-iframe" src="${GAS_URL}" allow="camera; geolocation; microphone"
    style="position:fixed;top:0;left:0;bottom:0;right:0;width:100%;height:100%;border:none;margin:0;padding:0;overflow:hidden;z-index:999999;"></iframe>

  <script>
    // ============ SPLASH ============
    const iframe = document.getElementById('main-iframe');
    const splash = document.getElementById('splash');
    iframe.onload = function(){
      iframe.classList.add('loaded');
      splash.style.opacity='0';
      setTimeout(()=>splash.style.display='none',500);
    };
    setTimeout(()=>{splash.style.opacity='0';setTimeout(()=>splash.style.display='none',500);iframe.classList.add('loaded');},8000);

    // ============ SERVICE WORKER + PUSH ============
    const VAPID_KEY = '${VAPID_PUBLIC}';
    let swRegistration = null;
    let pendingDni = null;

    async function initPush(){
      if(!('serviceWorker' in navigator) || !('PushManager' in window)) {
        console.log('Push no soportado en este navegador');
        return;
      }
      try {
        swRegistration = await navigator.serviceWorker.register('/sw.js', { scope: '/' });
        console.log('SW registrado:', swRegistration.scope);

        // Escuchar mensajes del iframe para suscribir/desuscribir
        window.addEventListener('message', async (e) => {
          if (!e.data || !e.data.type) return;

          if (e.data.type === 'PUSH_SUBSCRIBE') {
            pendingDni = e.data.dni;
            // Verificar si ya tenemos permiso
            if (Notification.permission === 'granted') {
              await doSubscribe(e.data.dni);
            } else if (Notification.permission !== 'denied') {
              // Mostrar botón para que el usuario active con gesto
              showNotificationBanner();
            }
          }
          if (e.data.type === 'PUSH_UNSUBSCRIBE') {
            await unsubscribePush(e.data.dni);
          }
        });

        // Si ya tiene permiso, ocultar banner
        if (Notification.permission === 'granted' && pendingDni) {
          hideNotificationBanner();
        }
      } catch(err) {
        console.error('Error registrando SW:', err);
      }
    }

    function showNotificationBanner() {
      const banner = document.getElementById('push-banner');
      if (banner) banner.style.display = 'flex';
    }

    function hideNotificationBanner() {
      const banner = document.getElementById('push-banner');
      if (banner) banner.style.display = 'none';
    }

    // Este se llama desde el CLICK del botón (gesto del usuario)
    async function activarNotificaciones() {
      hideNotificationBanner();
      try {
        const permission = await Notification.requestPermission();
        if (permission === 'granted') {
          if (pendingDni) await doSubscribe(pendingDni);
        } else {
          console.log('Permiso de notificación denegado por el usuario');
        }
      } catch(err) {
        console.error('Error solicitando permiso:', err);
      }
    }

    async function doSubscribe(dni) {
      if (!swRegistration) return;
      try {
        let sub = await swRegistration.pushManager.getSubscription();
        if (!sub) {
          const key = urlBase64ToUint8Array(VAPID_KEY);
          sub = await swRegistration.pushManager.subscribe({
            userVisibleOnly: true,
            applicationServerKey: key
          });
        }

        const resp = await fetch('/api/push/subscribe', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ dni: dni, subscription: sub.toJSON() })
        });
        const result = await resp.json();
        console.log('Suscripción push guardada:', result);

        iframe.contentWindow.postMessage({ type: 'PUSH_SUBSCRIBED', ok: true }, '*');
      } catch(err) {
        console.error('Error suscribiendo push:', err);
        iframe.contentWindow.postMessage({ type: 'PUSH_SUBSCRIBED', ok: false, error: err.message }, '*');
      }
    }

    async function unsubscribePush(dni) {
      if (!swRegistration) return;
      try {
        const sub = await swRegistration.pushManager.getSubscription();
        if (sub) {
          await fetch('/api/push/unsubscribe', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ dni: dni, endpoint: sub.endpoint })
          });
          await sub.unsubscribe();
        }
        iframe.contentWindow.postMessage({ type: 'PUSH_UNSUBSCRIBED', ok: true }, '*');
      } catch(err) {
        console.error('Error desuscribiendo push:', err);
      }
    }

    function urlBase64ToUint8Array(base64String) {
      const padding = '='.repeat((4 - base64String.length % 4) % 4);
      const base64 = (base64String + padding).replace(/-/g, '+').replace(/_/g, '/');
      const raw = atob(base64);
      const arr = new Uint8Array(raw.length);
      for(let i=0;i<raw.length;i++) arr[i]=raw.charCodeAt(i);
      return arr;
    }

    // Iniciar push al cargar
    initPush();
  </script>
</body>
</html>`;

  return new Response(html, {
    headers: { 'Content-Type': 'text/html;charset=UTF-8' }
  });
}

// ============================================================
//  HELPERS
// ============================================================
function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json', ...corsHeaders() }
  });
}

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization'
  };
}
