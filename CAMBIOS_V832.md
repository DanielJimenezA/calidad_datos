# v8.3.2 — Tooltips persistentes y sin interferencia del cursor

- Al pasar el mouse sobre un icono ⓘ, la explicación se abre y **permanece visible** aunque el cursor abandone el icono.
- Se coloca preferentemente a un costado, evitando cubrir el icono y el puntero; en pantallas estrechas se coloca arriba o abajo.
- El tooltip no intercepta clics ni movimiento del mouse (`pointer-events: none`).
- Se cierra al hacer clic fuera, presionar Escape o desplazarse; otro icono abre su explicación.
- Clic en el mismo icono alterna abrir/cerrar. Compatible con teclado y pantallas táctiles.
- No se modifica el procesamiento, las reglas ni los JSON.

## Instalación rápida

Sustituye `js/tooltips.js` y `css/styles.css` en tu repositorio por los archivos de este paquete. Si el navegador conserva el anterior, recarga con Ctrl+F5.
