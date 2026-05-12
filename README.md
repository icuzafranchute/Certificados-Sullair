# 📄 Generador Universal de Certificados — Sullair Argentina

App web para generar automáticamente certificados de servicio a partir del listado mensual de vouchers. Pensada para ser usada por todas las sucursales de Sullair Argentina.

---

## 🚀 ¿Qué hace?

- Lee el **Excel de vouchers** que llega mensualmente de facturación
- Genera certificados en **Excel** (una hoja por cert) o **PDF** (un archivo por cert en ZIP)
- Clasifica el equipo por **Unidad de Negocio** (columna U.N.) — más confiable que el prefijo del interno
- Soporta internos estándar (E01-E06, A01-A06), viejos (3xxx, 4xxx) y sin regla (por U.N.)
- Período con **mes y año** completo: "Julio 2025"
- Opción de **ocultar precios** para clientes que no deben verlos
- Nombre de archivo: `Certificación - CLIENTE - Período.xlsx/zip`
- Cada PDF nombrado: `Interno-Empresa-Periodo.pdf`

---

## 📁 Estructura del repositorio

```
├── app.py                  ← App principal Streamlit
├── requirements.txt        ← Dependencias Python
├── Dockerfile              ← Deploy en Railway (24/7)
├── assets/
│   └── logo_sullair.jpg    ← Logo Sullair Argentina
└── README.md
```

---

## 👤 Sistema de usuarios

Cada persona crea su propia cuenta desde la app (tab "Crear cuenta") con:
Nombre, Apellido, Sucursal, Email y Contraseña (mínimo 6 caracteres).

Usuario auto-generado: primera letra del nombre + apellido (ej: `JICUZA`).

**Usuario inicial:** `JICUZA` / `sullair2026`

---

## 📊 Columnas del Excel de vouchers

| Columna | Campo |
|---------|-------|
| U.N.    | Unidad de Negocio (1=Compresor, 2=Generador, 3=Plataforma, 4=Manipulador, 5=Luminaria, 6=Mov.Suelo) |
| Nombre  | Cliente |
| Artículo | Modelo |
| Interno | N° interno |
| Cnt Días | Días |
| Inicio Período | Fecha desde |
| Fin Período | Fecha hasta |
| Pcio Unitario | Valor diario |
| Total | Valor total |
| Moneda | Moneda |

---

## 🏢 Sucursales

Neuquén · Buenos Aires · Mendoza · Bahía Blanca · Comodoro Rivadavia · Salta · Tucumán · San Juan · Córdoba · Mar del Plata · Olavarría

---

## 🚂 Deploy en Railway (recomendado — 24/7)

1. [railway.app](https://railway.app) → New Project → Deploy from GitHub repo
2. Seleccionar este repositorio
3. Railway detecta el `Dockerfile` automáticamente
4. Settings → Networking → Generate Domain → URL fija permanente

---

## 🛠️ Correr localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

*Sullair Argentina S.A. · Generador Universal de Certificados · v5.0*
