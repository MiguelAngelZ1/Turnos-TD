# Turnos TD

Página web estática con diseño de app nativa para administrar turnos del personal. Los datos se guardan en el navegador (localStorage).

## Cómo usar

1. Abre `index.html` en tu navegador (doble clic o desde un servidor local).
2. En **Personal** agrega el equipo (nombre y estado activo/inactivo).
3. En **Turnos** verás el mes actual; toca cualquier día para asignar o cambiar la persona.
4. **Regla:** quien tiene turno el **viernes** tiene también **sábado y domingo** (se rellenan solos).
5. Al eliminar o marcar alguien como inactivo, sus turnos se **reasignan automáticamente** al resto.
6. En **Estadísticas** ves el conteo por persona: total, días laborales y fines de semana del mes.
7. En **Exportar** descargas un Excel del mes con colores y formato listo para imprimir o compartir.

## Requisitos

- Navegador moderno (Chrome, Edge, Firefox, Safari).
- Para exportar Excel se usa ExcelJS desde CDN; hace falta conexión a internet la primera vez (o incluir la librería en el proyecto).

## Estructura

- `index.html` — Estructura y modales
- `styles.css` — Estilos tipo app móvil (safe area, navegación inferior)
- `app.js` — Lógica de personal, turnos, reasignación y exportación
