# Utopía Comandas

App de pedidos, barra, caja y números de Utopía Coffee Club. Se instala como icono en el móvil (Android) y en el iPad/iPhone desde el navegador, sin tiendas de apps.

## Cómo está hecha

- **Una sola página web** (`index.html`) que funciona como app: pedidos, barra, caja, gastos, datos y ajustes.
- **Base de datos en Supabase** (`supabase/schema.sql`): una tabla `docs` con todo (carta, pedidos, cierres, gastos…) y cambios en tiempo real entre dispositivos.
- **Solo entra el equipo**: la tabla `equipo` guarda los correos autorizados. Cualquier otro correo no ve ni guarda nada.
- **Sin cobertura**: `sw.js` guarda la app en el dispositivo para que abra sin internet, y los pedidos hechos sin conexión se envían solos al volver.

## Puesta en marcha

1. En Supabase, ejecutar `supabase/schema.sql` y luego `supabase/seed.sql` (carta y suministros).
2. Añadir los correos del equipo: `insert into equipo (email) values ('correo1'), ('correo2');`
3. Poner la URL del proyecto y la clave `anon` en `config.js`.
4. Publicar la carpeta en GitHub Pages (o cualquier hosting de archivos estáticos con https).
5. En cada dispositivo: abrir el enlace, «Añadir a pantalla de inicio», y entrar con el correo («Es mi primera vez» la primera vez).

## Importante

Esta app organiza pedidos, caja y números internos. **No emite tickets ni facturas**: para eso seguid con un sistema homologado (normativa Veri*factu).
