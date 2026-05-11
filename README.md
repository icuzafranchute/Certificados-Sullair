# 📄 Generador Universal de Certificados — Sullair Argentina

App web para generar automáticamente certificados de servicio a partir del listado mensual de vouchers. Pensada para ser usada por todas las sucursales de Sullair Argentina.

---

## 🚀 ¿Qué hace?

- Lee el **Excel de vouchers** que llega mensualmente de facturación
- Genera un **archivo Excel** con un certificado por hoja, fiel al modelo oficial de Sullair
- Clasifica automáticamente el tipo de equipo (Generador, Compresor, Plataforma, Manipulador, Luminaria)
- Nombra el archivo como: `Certificación - CLIENTE - Período.xlsx`
- Permite que cada sucursal genere sus propios certificados desde cualquier navegador

---

## 📁 Estructura del repositorio

```
├── app.py                  ← App principal Streamlit
├── requirements.txt        ← Dependencias Python
├── Dockerfile              ← Para deploy en Railway (sin interrupciones)
├── assets/
│   └── logo_sullair.jpg    ← Logo Sullair Argentina
└── README.md
```

---

## 👤 Sistema de usuarios

Cada persona crea su propia cuenta desde la app con:
- Nombre y Apellido
- Sucursal
- Email
- Contraseña (mínimo 6 caracteres)

El usuario se genera automáticamente (ej: primera letra del nombre + apellido = `JICUZA`).

Los usuarios se guardan en `usuarios.json` (creado automáticamente al iniciar la app).

**Usuario inicial de administración:**
- Usuario: `JICUZA`
- Contraseña: `sullair2026`

---

## 📊 Fuente de datos — Excel de vouchers

La app lee el Excel mensual que envía facturación. La estructura esperada es:

| Fila | Contenido |
|------|-----------|
| 1    | Período (ej: "Julio") |
| 2    | Encabezados de columnas |
| 3+   | Datos de vouchers |

Columnas que usa la app:

| Columna | Campo |
|---------|-------|
| Nombre  | Cliente |
| Artículo | Modelo del equipo |
| Interno | N° interno del equipo |
| Cnt Días | Cantidad de días |
| Inicio Período | Fecha desde |
| Fin Período | Fecha hasta |
| Pcio Unitario | Valor diario |
| Total | Valor total |
| Moneda | Moneda (U$S, ARS, etc.) |

---

## 🏢 Sucursales disponibles

- Neuquén · Buenos Aires · Mendoza · Bahía Blanca · Comodoro Rivadavia
- Salta · Tucumán · San Juan · Córdoba · Mar del Plata · Olavarría

---

## 🛠️ Deploy en Railway (recomendado — 24/7 sin interrupciones)

1. Crear cuenta en [railway.app](https://railway.app)
2. New Project → **Deploy from GitHub repo**
3. Seleccionar este repositorio
4. Railway detecta el `Dockerfile` automáticamente
5. En 2-3 minutos la app está disponible en una URL fija permanente

---

## 🛠️ Deploy en Streamlit Cloud (alternativa gratuita)

1. Ir a [share.streamlit.io](https://share.streamlit.io)
2. New app → seleccionar este repositorio → Main file: `app.py`

> ⚠️ En plan gratuito puede pausar la app si no se usa por un tiempo.

---

## 🔧 Correr localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## 📝 Notas técnicas

- El certificado replica el modelo oficial de Sullair con columnas y alturas exactas del modelo
- El logo se posiciona con anchor `I4` (253×53px) según el XML del archivo modelo
- Los bordes del certificado usan estilo `medium` replicando el modelo celda por celda
- La clasificación de equipo se hace por artículo primero, luego por prefijo del interno
- Internos `3xxx` = Plataformas, `4xxx` = Manipuladores (numeración anterior a la estandarización)

---

*Sullair Argentina S.A. · Generador Universal de Certificados · v4.0*
