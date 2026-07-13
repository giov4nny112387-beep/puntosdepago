# Programaciones Puntos de Pago — App Web

Aplicación web (HTML + JS + CSS) para visualizar y editar las mallas de turnos
de los puntos de pago. Se conecta a Google Sheets mediante un Apps Script.

## Contenido de esta carpeta (todo lo necesario para funcionar)

```
index.html            La aplicación (página principal)
css/app.css           Estilos
js/app.js             Toda la lógica
libs/chart.min.js           Gráficos
libs/xlsx.full.min.js       Lectura de Excel
libs/html2pdf.bundle.min.js Exportar a PDF
```

No hace falta nada más: la app es autónoma (todas las librerías van incluidas
en `libs/`, no se descarga nada de internet).

## Publicar en GitHub Pages

1. Crea un repositorio en GitHub y sube **todo el contenido de esta carpeta**
   (respetando las subcarpetas `css/`, `js/`, `libs/`).
2. En el repo: **Settings → Pages → Deploy from a branch → main / (root)**.
3. Espera 1–2 min. La app quedará en:
   `https://TU_USUARIO.github.io/TU_REPO/`
   (como el archivo se llama `index.html`, la URL abre directo, sin nombre extra).

## Configuración

- La URL del Apps Script está en `index.html`, en `SHEETS_CONFIG.APPS_SCRIPT_URL`.
- Login de la demo: usuario `administrador`, clave `ADMIN1`.

## Seguridad (importante)

- El login es de **cortesía**, no seguridad real (está en el código visible).
- La URL del Apps Script queda visible al publicar; si está como "cualquiera
  anónimo", cualquiera con el enlace puede leer/escribir ese Google Sheets.
- **NUNCA** subas a este repo la llave `credenciales_service_account.json` ni
  archivos de ventas/planta (`.xlsx`): son datos sensibles de la compañía.
