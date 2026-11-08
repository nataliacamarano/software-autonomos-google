# Normas de trabajo en este proyecto

Programa de Google Apps Script para gestionar el negocio de autónomos de Álvaro (facturas, presupuestos, gastos, paneles), vinculado a una hoja de Google Sheets.

- Repositorio: https://github.com/nataliacamarano/software-autonomos-google (rama `main`).
- Script de Google (scriptId): `1wWbcEDAfecpUDyP7PigBFaLR-tX3qB4J4KoXDPFyYiDp_pLCFtGX5jC2`.
- La usuaria (Natalia) no es técnica: háblale en español y con palabras sencillas.

## Normas obligatorias

1. **En PowerShell usa siempre `clasp.cmd`, nunca `clasp`.** La política de ejecución bloquea `clasp.ps1`. No cambies esa política.
2. **NUNCA hagas `git push` ni `clasp.cmd push` sin la confirmación explícita de Natalia en ese momento.** Que lo aprobara antes no vale para la siguiente vez.
3. **Antes de un `clasp.cmd push`, enséñale `clasp.cmd status`** (la lista de archivos que se van a subir).
4. **Antes de un `clasp.cmd push`, comprueba que nadie haya editado en el editor web de Google:**
   - Clona el script en una carpeta temporal (fuera del repo): `clasp.cmd clone <scriptId>`.
   - Compara esos archivos con los locales (ten en cuenta que el clon puede traer `.js` en vez de `.gs`; compara el contenido).
   - Si hay diferencias, avisa a Natalia antes de hacer nada más.
5. **No edites ni borres `.clasp.json`.** Está en `.gitignore`, no está en GitHub y solo existe en esta carpeta.
6. **Commits en español**, describiendo el cambio.

## Al empezar una sesión

- `git pull` (rama `main`) y `clasp.cmd status` para comprobar que todo está en orden.
- clasp tiene sesión iniciada como alvaro.mantenimiento.ibiza@gmail.com. Natalia decidió dejarlo así.
