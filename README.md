# Automatización de Movimientos Contables

## Descripción

Este script automatiza la actualización de los movimientos contables de pérdidas y ganancias de una empresa. 

Compara los movimientos del archivo `InputPL.xlsx` con los del archivo `Mayor_TSCFO.xlsx` y realiza las siguientes tareas:

- Detecta movimientos nuevos que no estén en `InputPL.xlsx`.
- Evita duplicados mediante un ID interno generado a partir de los campos clave.
- Añade los movimientos nuevos manteniendo el formato original del archivo.
- Asigna automáticamente un **Tipo de gasto** basándose en movimientos históricos.
- Calcula un grado de **Confianza (%)** para cada asignación.
- Marca como `REVISAR` los movimientos cuya confianza esté por debajo del umbral definido.

--- 

## Tecnologías utilizadas

- Python 3.13
- pandas 
- openpyxl
- difflib (SequenceMatcher)
- Poetry (opcional, también se puede instalar con pip)

---
## Estructura del proyecto
```
automatizar_movimentos/
├── main.py
├── InputPL.xlsx
├── Mayor_TSCFO.xlsx
├── InputPL_actualizado.xlsx (generado)
├── poetry.lock
├── pyproject.toml
└── README.md
```
---

## Instalación y ejecución

### Instalar dependencias
```bash
  poetry install
```
### Ejecutar
```bash
  poetry run python main.py
```
Tras la ejecución, se generará el archivo:
```
InputPL_actualizado.xlsx
```
con los nuevos movimientos añadidos y clasificados.

---

Lógica de funcionamiento
1. Se cargan ambos archivos Excel.
2. Se genera un ID interno para detectar duplicados.
3. Se identifican movimientos nuevos.
4. 	Se clasifica cada nuevo movimiento:
* Si la similitud supera el umbral → se asigna automáticamente el Tipo de gasto.
* Si no lo supera → se marca como REVISAR.
5. Se calcula la columna Confianza (valor numérico entre 0 y 100).
6. Se insertan las nuevas filas antes de END, manteniendo el formato original.
7. Se muestra un resumen de ejecución en consola.


### Configuración
El umbral de similitud puede modificarse en:
```
UMBRAL_SIMILITUD = 0.75
```
