# Home Banking

Dashboard personal de finanzas: centraliza movimientos bancarios de distintos bancos y efectivo, en pesos y dólares, con un solo login simple.

Construido para un usuario real que necesitaba reemplazar planillas de Excel dispersas por un único lugar donde ver, cargar y corregir sus movimientos.

**🔗 [Demo pública](https://home-banking-demo-441756806758.southamerica-east1.run.app)** — usuario `demo`, contraseña `demo1234`. Datos de ejemplo, no reales.

## Qué hace

- **Importa extractos bancarios** (Excel) de distintos bancos (Galicia, Macro, Mercado Pago) detectando automáticamente el formato, la moneda y la cuenta de origen de cada archivo.
- **Carga manual de efectivo**, con la misma clasificación automática ARS/USD.
- **Dashboard multi-moneda**: pesos y dólares en paneles separados, con filtros por período y cuenta.
- **Saldo actual real por cuenta**: permite cargar un saldo inicial para que el balance mostrado coincida con la plata real, no solo con lo cargado en la app.
- **Tabla editable**: corrige fecha, descripción, monto o tipo de un movimiento ya guardado, sin tener que borrarlo y recargarlo.
- **Papelera con restauración**: los borrados son reversibles (soft delete) y se purgan solos después de 30 días.
- **Reportes en PDF** mensuales, independientes de los filtros del dashboard.
- **Login simple** con usuario/contraseña y sesión persistente (cookie), pensado para un usuario no técnico.
- **Instalable como app** en el celular (PWA) — ícono propio, sin la barra del navegador.

## Stack

- **Streamlit** — UI y lógica de la app
- **PostgreSQL** ([Neon](https://neon.tech), serverless) — persistencia
- **Pandas** — parseo y normalización de extractos bancarios
- **Plotly** — gráficos del dashboard
- **fpdf2** — generación de reportes PDF
- **streamlit-authenticator** — login con cookie
- **Docker + Google Cloud Run** — deploy, con CI/CD automático desde GitHub

## Correrlo localmente

Para levantarlo local hace falta tu propia base Postgres (arranca vacía, sin datos de ejemplo) — para ver la app andando con datos, la demo pública de arriba es más directa.

```bash
pip install -r requirements.txt
```

Crear `.streamlit/secrets.toml`:

```toml
DATABASE_URL = "postgresql://usuario:password@host/db"
APP_USERNAME = "admin"
APP_PASSWORD = "tu-contraseña"
APP_COOKIE_KEY = "cualquier-string-largo-random"
```

```bash
streamlit run streamlit_app.py
```

## Tests

```bash
pytest test_app.py -v
```

Cobertura de detección de moneda, parseo de extractos, deduplicación de movimientos, edición/borrado, y flujos completos de UI con `streamlit.testing.v1.AppTest`.

## Deploy

Imagen Docker corriendo en Google Cloud Run, con despliegue automático en cada push a `main` vía Cloud Build.
