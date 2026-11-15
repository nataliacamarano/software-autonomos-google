# Álvaro Gestión — app instalable

"Puerta de entrada" que permite instalar el programa de Google (web app de Apps Script) como una app en el móvil, con icono propio y pantalla completa. No contiene datos ni claves: cada persona entra con su enlace personal (`?t=...`), igual que ahora.

## Archivos
- `index.html`: abre el programa de Google dentro de la app y recuerda el enlace personal.
- `manifest.webmanifest`, `sw.js`, `icon-*.png`, `apple-touch-icon.png`: lo necesario para que el móvil ofrezca "Instalar".
- `logo-original.png`: logo de Álvaro Mantenimiento Ibiza. Los iconos son un recorte de su emblema (tejado, llave y sol) sobre fondo blanco.

## Cómo entra cada persona
1. Instala la app desde la dirección general (sin clave).
2. La primera vez, pega el enlace personal que le mandó Álvaro (el que genera el programa al dar de alta al empleado). La app saca la clave y la recuerda.
3. Si el enlace deja de ser válido, la app lo olvida y pide uno nuevo.

## Ajustes (botón ⚙️ del programa)
- **PIN opcional:** se pide al abrir la app y al volver a ella pasado un minuto. Se guarda solo en el móvil, cifrado. 5 fallos seguidos bloquean 30 segundos.
- **Cerrar sesión:** la app olvida el enlace personal y el PIN; hay que pegar el enlace otra vez.
- Si alguien olvida el PIN: "Olvidé mi PIN" cierra la sesión y se pega el enlace de nuevo.

## Si cambia la dirección del programa de Google
Editar la constante `EXEC` en `index.html` (misma dirección `.../exec` de la implementación).

## Publicación
Este repositorio (`software-autonomos-google`) es la **única fuente de verdad** de la app. Se edita solo aquí, en `app-instalable/`.
- Al fusionar un cambio en `main`, un proceso automático (`.github/workflows/publicar-app.yml`) publica la carpeta en GitHub Pages: `https://nataliacamarano.github.io/software-autonomos-google/`.
- Esa misma dirección está en `URL_APP` de `WebApp.gs` (los enlaces personales la usan). Si cambia, hay que cambiar las dos cosas y hacer `clasp push`.
- `.claspignore` impide que `clasp push` suba esta carpeta a Google.
- El repositorio antiguo `alvaro-gestion-app` ya no se debe editar: solo sigue sirviendo la dirección vieja para quien ya tenga la app instalada de antes.



