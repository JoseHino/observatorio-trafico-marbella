# Balances de criminalidad (PDF del Ministerio del Interior)

Deja aquí los **balances trimestrales de criminalidad** en PDF (los del **4º
trimestre** dan el año completo). El colector (`scripts/collect.py`) los lee
automáticamente, extrae la tabla del *Municipio de Marbella* y actualiza
`data/criminalidad.json`.

- Cada balance del 4º trimestre aporta ese año **+ el anterior ya consolidado**.
- Para cubrir 2017-2025 bastan los del 4º trimestre de 2018, 2020, 2022, 2024 y 2025.
- Descárgalos desde el portal oficial: https://estadisticasdecriminalidad.ses.mir.es/
  (sección *Balances trimestrales*). El Ministerio bloquea las descargas
  automáticas desde servidores, por eso se guardan aquí a mano.

Nombra el fichero como quieras (p. ej. `balance_2022.pdf`); el año se detecta
del contenido.
