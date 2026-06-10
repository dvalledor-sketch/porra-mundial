# Auditoría de seguridad — Porra Mundial 2026 / Supabase

Fecha: 9 jun 2026 · Proyecto Supabase: `iwicyguwsobdbsusxydx`

## Resumen ejecutivo

La aplicación de la porra está **bien construida**: las claves secretas no están en el repo,
las escrituras pasan por funciones RPC validadas, y se respeta la hora límite.

**Pero** la porra comparte el mismo proyecto Supabase que una base de datos de **producción
con datos personales de clientes** (telco: `customers`, `invoices`, `payments`, `cdrs`,
`sim_cards`…). Varias de esas tablas tienen políticas RLS **abiertas al rol anónimo
(`USING (true)`)**, y la clave pública de ese rol está publicada en el sitio web de la porra
(GitHub Pages) y en el repositorio. Resultado: **cualquiera con esa clave pública puede leer
—y en algunos casos insertar— datos sensibles de clientes.**

Esto es un riesgo **crítico** y conviene actuar hoy. No es culpa del diseño de la porra:
son políticas heredadas del otro proyecto que vive en la misma base de datos.

---

## 🔴 CRÍTICO — Datos de producción expuestos al rol anónimo

La clave publicable `sb_publishable_Q8aBMLOXDNaqZgp-qmoZwA_5kj79yWp` aparece en:
`dashboard/pronostico.html`, `dashboard/index.html`/`dashboard.html` y el repo público.

En Supabase, una clave publicable **no es secreta por diseño**: la seguridad real la dan las
políticas RLS. El problema es que estas tablas tienen `SELECT ... TO anon USING (true)`
(lectura total) **e** `INSERT ... TO anon WITH CHECK (true)` (inserción total):

| Tabla | Lectura anon | Inserción anon |
|---|---|---|
| `customers` | ✅ abierta | ✅ abierta |
| `invoices`, `invoice_lines` | ✅ abierta | ✅ abierta |
| `payments`, `payment_methods` | ✅ abierta | ✅ abierta |
| `sim_cards`, `subscriptions` | ✅ abierta | ✅ abierta |
| `cdrs` (registros de llamadas) | ✅ abierta | ✅ abierta |
| `consent_records` | ✅ abierta | — |
| `portability_requests` | ✅ abierta | — |
| `support_tickets`, `ticket_messages` | ✅ abierta | — |

Impacto: filtración masiva de PII (nombres, emails, facturación, pagos, líneas, llamadas),
posible relevancia RGPD, y manipulación/inyección de registros falsos.

**Nota:** se confirmó por la *definición* de las políticas, sin extraer ningún dato real.

### Remediación (revisar y aplicar con cuidado — afecta al otro proyecto)

Estas políticas pertenecen a la app telco, no a la porra. Antes de aplicar, confirma que esa
app **no** depende del acceso anónimo (debería usar el rol `service_role` o usuarios
autenticados). SQL sugerido:

```sql
-- Quitar lectura/escritura anónima de las tablas sensibles
drop policy if exists allow_read_customers            on public.customers;
drop policy if exists allow_read_invoices             on public.invoices;
drop policy if exists allow_read_invoice_lines        on public.invoice_lines;
drop policy if exists allow_read_payments             on public.payments;
drop policy if exists allow_read_payment_methods      on public.payment_methods;
drop policy if exists allow_read_sim_cards            on public.sim_cards;
drop policy if exists allow_read_subscriptions        on public.subscriptions;
drop policy if exists allow_read_cdrs                 on public.cdrs;
drop policy if exists allow_read_consent_records      on public.consent_records;
drop policy if exists allow_read_portability_requests on public.portability_requests;
drop policy if exists allow_read_support_tickets      on public.support_tickets;
drop policy if exists allow_read_ticket_messages      on public.ticket_messages;
-- (repetir el patrón para las políticas INSERT equivalentes)
```

Si la app telco necesita esos datos en cliente, reemplaza `USING (true)` por una condición
ligada al usuario, p. ej. `USING (customer_id = auth.uid())`, como ya hacen correctamente
las tablas `calls` y `profiles`.

### Mejor aún: separar proyectos
Lo ideal es que la porra (datos triviales y públicos) **no comparta proyecto** con producción.
Mover la porra a su propio proyecto Supabase elimina de raíz que su clave pública toque datos
de clientes.

---

## 🟡 MEDIO — Emails de participantes legibles por anon

`porra_participantes` tiene `SELECT TO anon USING (true)` sobre **todas** las columnas,
incluida `email`. El dashboard solo pide `alias_norm, alias, updated_at`, pero cualquiera con
la clave pública puede pedir también `email` y listar los correos de los participantes.

### Remediación
Exponer solo las columnas necesarias mediante una **vista** y dar SELECT anon a la vista, no a
la tabla:

```sql
revoke select on public.porra_participantes from anon;
create or replace view public.porra_participantes_pub as
  select alias_norm, alias, updated_at from public.porra_participantes;
grant select on public.porra_participantes_pub to anon;
```
Y en el cliente/JS leer de `porra_participantes_pub` en vez de `porra_participantes`.
(Las funciones RPC `porra_*` son `SECURITY DEFINER`, así que siguen viendo la tabla completa.)

---

## 🟢 BIEN — Lo que ya está correcto

- **`service_role` no está en el repo**: solo en GitHub Secrets (`SUPABASE_SERVICE_KEY`),
  usado en los workflows. Correcto.
- **Escrituras de la porra vía RPC**: el cliente solo llama a `porra_cargar` y `porra_guardar`
  (`/rest/v1/rpc/`), no escribe en tablas directamente.
- **`porra_guardar` es robusta**: `SECURITY DEFINER` con `SET search_path = public`, valida
  formato de email, **respeta la hora límite** (`now() > limite → 'cerrado'`), exige que el
  email coincida si el alias ya existe (evita suplantación), y usa JSONB parametrizado
  (sin inyección SQL).
- **RLS activado** en todas las tablas del esquema `public`.

---

## Prioridad de acción

1. **HOY** — Cerrar lectura/inserción anónima en las tablas de producción (sección crítica).
2. **Esta semana** — Ocultar `email` en `porra_participantes` tras una vista.
3. **Cuando puedas** — Rotar la clave publicable y, a medio plazo, mover la porra a su propio
   proyecto Supabase.
