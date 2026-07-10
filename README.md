# Observatorio de Tráfico de Marbella (datos DGT)

Recolección **automática** y visualización de los datos municipales de tráfico de la
Dirección General de Tráfico (DGT) para **Marbella**, con serie histórica completa y
comparativas frente a Málaga capital, la provincia de Málaga, Andalucía y España.

**Dashboard en vivo:** https://josehino.github.io/observatorio-trafico-marbella/

## ¿Qué datos incluye?

Toma los tres conjuntos oficiales de *DGT en cifras · Información municipal* (un Excel
por año con todos los municipios de España) y extrae Marbella + agregados de su entorno:

| Dataset | Contenido | Cobertura |
|---|---|---|
| **General** | Población, censo de conductores, parque por tipo/antigüedad/combustible, distintivos ambientales | 2015–2025 |
| **Siniestralidad** | Accidentes con víctimas, fallecidos y heridos por tipo de vehículo y peatones | 2015–2024 |
| **Sanciones y puntos** | Sanciones con puntos y puntos detraídos por tipo de infracción | 2015–2023 |

Los años avanzan solos según la DGT va publicando (la siniestralidad y las sanciones
salen con ~1 año de retraso).

## Salidas

- `docs/index.html` — dashboard web (GitHub Pages).
- `data/trafico_marbella.json` — datos estructurados (Marbella + 4 ámbitos de comparación).
- `Observatorio_Trafico_Marbella_DGT.xlsx` — Excel con portada, series históricas,
  comparativa e indicadores derivados (motorización, % parque limpio, tasas de siniestralidad…).

## Uso local

```bash
pip install -r requirements.txt
python scripts/collect.py      # descarga y procesa todos los años
python scripts/build_excel.py  # genera el Excel enriquecido
```

## Automatización

El workflow `.github/workflows/update.yml` se ejecuta el día 5 de cada mes (y a mano
desde la pestaña *Actions*): descarga los ficheros de la DGT, regenera el JSON y el
Excel, y hace commit solo si hay cambios. El dashboard se actualiza automáticamente.

## Fuente

Dirección General de Tráfico — [DGT en cifras](https://www.dgt.es/menusecundario/dgt-en-cifras/)
(Tema: *Información municipal*). Datos abiertos, descarga HTTP directa sin API key.
