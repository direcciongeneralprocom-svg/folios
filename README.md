# Generador de Formato de Operaciones

Página web estática (HTML + JS) que lee la factura (XML + PDF), te permite elegir el cliente
(Luis Pape, Berna, Estela H, Luis Avila) y genera automáticamente el Excel de "Formato de
Operaciones" ya rellenado con los datos de la factura, conservando las fórmulas, comisiones y
cuentas de retorno de cada plantilla.

Todo el procesamiento ocurre en el navegador del usuario (no hay backend ni se sube ninguna
factura a ningún servidor).

## Contenido

- `index.html` — la página.
- `app.js` — lógica: lectura del XML CFDI 4.0 y llenado del Excel (usa la librería ExcelJS desde CDN).
- `templates/template_pape.xlsx`, `template_estela.xlsx`, `template_avila.xlsx` — plantilla
  "FORMATO DE OPERACIONES" (FOPAS-12-01) con la comisión y cuentas de retorno de cada cliente.
- `templates/template_berna.xlsx` — plantilla "SOLICITUD DE FACTURACION" de Berna.

## Cómo funciona

1. Subes el **XML** de la factura (de ahí se leen folio, fecha, emisor, receptor, RFC, conceptos,
   importes, forma/método de pago). El **PDF** se adjunta solo como respaldo/consulta, no se lee
   su contenido.
2. Eliges el cliente. Cada cliente tiene fija su comisión y sus cuentas de retorno (igual que en
   los archivos que ya manejabas); la "EMPRESA" emisora se toma de la factura, no del cliente.
3. El formulario te deja completar dirección del cliente (calle/colonia/ciudad/CP) y fechas de
   pago/devolución, porque el XML del SAT no las incluye.
4. Al dar "Generar y descargar Excel" se descarga el archivo `.xlsx` ya armado.

### Importante / limitaciones a revisar

- Si una factura tiene **más de un concepto**, la herramienta los junta en un solo renglón
  (suma importes, concatena descripciones) para no romper las fórmulas de la plantilla. Si
  necesitas que cada concepto vaya en su propio renglón, dímelo y lo ajustamos.
- Las **cuentas de retorno / beneficiarios** quedaron tal como estaban en tus archivos de ejemplo
  (de referencia). Revísalas antes de usar el Excel para pagos reales.
- El **régimen fiscal / uso de CFDI / forma de pago** se traducen a texto con el mismo catálogo
  que ya usabas en la hoja "Datos". Si el SAT agrega un código nuevo que no está en la lista,
  se mostrará el código tal cual (revisar manualmente).

## Publicar en GitHub Pages

1. Crea un repositorio nuevo en GitHub (puede ser privado), por ejemplo `formato-operaciones`.
2. Sube estos 3 elementos a la raíz del repo: `index.html`, `app.js`, y la carpeta `templates/`
   completa (con los 4 `.xlsx`).
   - Más fácil: arrastra los archivos en la página web de GitHub ("Add file" → "Upload files"),
     o usa GitHub Desktop / `git push` si ya usas git.
3. Ve a **Settings → Pages** del repositorio.
4. En "Branch" selecciona `main` (o la rama donde subiste los archivos) y carpeta `/ (root)`.
   Guarda.
5. GitHub te dará una URL tipo `https://tu-usuario.github.io/formato-operaciones/` — ábrela en
   unos minutos y la página debe funcionar igual que en local.

No necesitas servidor, base de datos ni build: es un sitio 100% estático.
