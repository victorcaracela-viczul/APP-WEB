addEventListener('fetch', event => {
  event.respondWith(handleRequest(event.request))
})
async function handleRequest(request) {
  const googleAppScriptUrl = 'https://script.google.com/macros/s/AKfycbwJrer0KO6jEd9HFso-AKzARyzlVdRrblJzm1H2i2ylWCbsCS9XzLGAfuQio2EPMzg/exec';
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
        <title>ADECCO - ISOS</title>

        <!-- 🎨 FUENTE MODERNA DE GOOGLE -->
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@600;700&family=Orbitron:wght@700&display=swap" rel="stylesheet">

        <!-- 🖥️ ICONO DE PESTAÑA -->
        <link rel="icon" type="image/png" href="https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o">

        <!-- 📱 ICONOS PARA iOS -->
        <link rel="apple-touch-icon" sizes="180x180" href="https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o">
        <link rel="apple-touch-icon" sizes="152x152" href="https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o">
        <link rel="apple-touch-icon" sizes="144x144" href="https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o">
        <link rel="apple-touch-icon" sizes="120x120" href="https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o">

        <!-- 🤖 ANDROID ICONOS -->
        <link rel="icon" type="image/png" sizes="512x512" href="https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o">
        <link rel="icon" type="image/png" sizes="192x192" href="https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o">
        <link rel="icon" type="image/png" sizes="96x96" href="https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o">
        <link rel="icon" type="image/png" sizes="32x32" href="https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o">

        <!-- 🎨 CONFIGURACIÓN PWA CON ORIENTACIÓN LIBRE -->
        <meta name="theme-color" content="#2c3e50">
        <meta name="apple-mobile-web-app-capable" content="yes">
        <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
        <meta name="apple-mobile-web-app-title" content="ADECCO - ISOS">
        <meta name="mobile-web-app-capable" content="yes">
        <meta name="application-name" content="ADECCO - ISOS">

        <!-- 🔓 Permitir rotación libre -->
        <meta name="screen-orientation" content="any">

        <!-- 🔗 MANIFEST CON ORIENTACIÓN "ANY" -->
        <link rel="manifest" href="data:application/json;base64,\${btoa(JSON.stringify({
          "name": "ADECCO - ISOS",
          "short_name": "ADECCO - ISOS",
          "description": "Sistema de Gestión Operacional SSOMA",
          "start_url": "/",
          "display": "standalone",
          "background_color": "#ffffff",
          "theme_color": "#2c3e50",
          "orientation": "any",
          "icons": [
            {
              "src": "https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o",
              "sizes": "72x72",
              "type": "image/png",
              "purpose": "any"
            },
            {
              "src": "https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o",
              "sizes": "96x96",
              "type": "image/png",
              "purpose": "any"
            },
            {
              "src": "https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o",
              "sizes": "128x128",
              "type": "image/png",
              "purpose": "any"
            },
            {
              "src": "https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o",
              "sizes": "144x144",
              "type": "image/png",
              "purpose": "any"
            },
            {
              "src": "https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o",
              "sizes": "152x152",
              "type": "image/png",
              "purpose": "any"
            },
            {
              "src": "https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o",
              "sizes": "192x192",
              "type": "image/png",
              "purpose": "any"
            },
            {
              "src": "https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o",
              "sizes": "384x384",
              "type": "image/png",
              "purpose": "any"
            },
            {
              "src": "https://lh3.googleusercontent.com/d/1qik5wQ9CWfURpqBP4LI3YzMO_PHEyt6o",
              "sizes": "512x512",
              "type": "image/png",
              "purpose": "any"
            }
          ]
        }))}">

        <style>
          body, html {
            margin: 0;
            padding: 0;
            height: 100vh;
            width: 100vw;
            overflow: hidden;
            background-color: #ffffff;
          }

          /* 🎬 SPLASH SCREEN */
          #splash {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: #ffffff;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            z-index: 9999;
            transition: opacity 0.5s ease-out;
          }

          /* 🎬 GIF CENTRADO */
          #splash img {
            width: 200px;
            height: auto;
            max-height: 200px;
          }

          /* ✨ TEXTO SUPER ANIMADO CON TIPOGRAFÍA MODERNA */
          #splash-text {
            font-family: 'Orbitron', 'Poppins', sans-serif;
            font-size: 22px;
            font-weight: 700;
            margin-top: 30px;
            text-align: center;
            letter-spacing: 4px;
            background: linear-gradient(90deg, #8b0000, #ff0000, #8b0000);
            background-size: 200% auto;
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            animation: gradient-shift 3s ease infinite, pulse-text 2s ease-in-out infinite;
            text-transform: uppercase;
            position: relative;
          }

          /* 🌈 Animación de gradiente */
          @keyframes gradient-shift {
            0% { background-position: 0% center; }
            50% { background-position: 100% center; }
            100% { background-position: 0% center; }
          }

          /* 💫 Animación de pulso sutil */
          @keyframes pulse-text {
            0%, 100% {
              transform: scale(1);
              filter: brightness(1);
            }
            50% {
              transform: scale(1.02);
              filter: brightness(1.2);
            }
          }

          /* ✨ Efecto de brillo que pasa */
          #splash-text::before {
            content: '';
            position: absolute;
            top: 0;
            left: -100%;
            width: 100%;
            height: 100%;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.8), transparent);
            animation: shine 3s infinite;
          }

          @keyframes shine {
            0% { left: -100%; }
            50%, 100% { left: 100%; }
          }

          /* 🔄 PUNTOS SUSPENSIVOS MEJORADOS */
          .dots {
            display: inline-block;
            font-family: 'Orbitron', monospace;
            font-weight: 700;
          }

          .dots::after {
            content: '';
            animation: dots 1.5s steps(4, end) infinite;
          }

          @keyframes dots {
            0% { content: ''; }
            25% { content: '.'; }
            50% { content: '..'; }
            75% { content: '...'; }
            100% { content: ''; }
          }

          iframe {
            width: 100%;
            height: 100%;
            border: none;
            opacity: 0;
            transition: opacity 0.5s ease-in;
          }

          iframe.loaded {
            opacity: 1;
          }

          /* 📱 RESPONSIVE */
          @media (max-width: 768px) {
            #splash-text {
              font-size: 18px;
              letter-spacing: 3px;
            }

            #splash img {
              width: 160px;
              max-height: 160px;
            }
          }
        </style>
      </head>
      <body>
        <!-- 🎬 SPLASH CON GIF Y TEXTO ANIMADO -->
        <div id="splash">
          <img src="https://lh3.googleusercontent.com/d/15B-wj4iw5B7RpDXQg2mrdxTAn-3kfeMa" alt="Cargando">
          <div id="splash-text">CARGANDO SISTEMA<span class="dots"></span></div>
        </div>

        <!-- 📱 TU IFRAME ORIGINAL -->
        <iframe
          id="main-iframe"
          src="\${googleAppScriptUrl}"
          allow="camera; geolocation; microphone"
          style="position:fixed; top:0; left:0; bottom:0; right:0; width:100%; height:100%; border:none; margin:0; padding:0; overflow:hidden; z-index:999999;">
        </iframe>

        <script>
          // 🚀 LÓGICA SIMPLE DE SPLASH
          const iframe = document.getElementById('main-iframe');
          const splash = document.getElementById('splash');

          // Ocultar splash cuando el iframe cargue
          iframe.onload = function() {
            iframe.classList.add('loaded');
            splash.style.opacity = '0';
            setTimeout(() => {
              splash.style.display = 'none';
            }, 500);
          };

          // Fallback: ocultar después de 8 segundos máximo
          setTimeout(() => {
            splash.style.opacity = '0';
            setTimeout(() => {
              splash.style.display = 'none';
            }, 500);
            iframe.classList.add('loaded');
          }, 8000);
        </script>
      </body>
    </html>
  `;
  return new Response(html, {
    headers: { 'Content-Type': 'text/html;charset=UTF-8' },
  });
}
