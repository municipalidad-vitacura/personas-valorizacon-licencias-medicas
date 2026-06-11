# App Web: Monto Devolución por Licencia

App en Flask para subir un Excel, seleccionar un área global y devolver el mismo archivo con cálculos de monto de devolución por licencia.

## Flujo

1. Subir archivo Excel (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`).
2. Seleccionar área global: `Municipal`, `Salud` o `Educación`.
3. Procesar y descargar archivo modificado.

## Formato de entrada requerido

La hoja activa del Excel debe contener exactamente estas columnas en la fila de encabezados:

- `rut`
- `fecha_inicio`
- `fecha_fin`

### Fechas

- Se acepta fecha real de Excel (tipo fecha).
- En texto, solo formato `dd-mm-yy`.

### RUT

- El RUT se envía al servicio sin dígito verificador, sin puntos y sin guion.
- Ejemplo: `19.829.424-1` -> `19829424`.
- Mapeo de dominio por área:
  - `Municipal` -> `11`
  - `Educación` -> `2`
  - `Salud` -> `3`

## Resultado del procesamiento

1. En la hoja original:
   - Se agrega columna `monto_devolucion`.
   - Se agregan columnas de validación:
     - `validador_licencia_valida`
     - `validador_tipo_licencia`
     - `validador_dias`
     - `validador_fecha_desde`
     - `validador_fecha_hasta`
     - `validador_lugar_reposo`
     - `validador_entidad`
     - `validador_observacion`
2. Se crea/reemplaza hoja `detalle_mensual` con columnas:
   - `fila_origen`
   - `id_licencia` (referencial, puede venir vacío)
   - `rut`
   - `area`
   - `periodo_inicio`
   - `periodo_fin`
   - `mes`
   - `remuneracion`
   - `dias_habiles`
   - `valor_dia`
   - `monto_mes`

## Cálculos

- El período `fecha_inicio` a `fecha_fin` se divide por meses calendario.
- `dias_habiles` se calcula como lunes-viernes excluyendo feriados nacionales de Chile.
- Se valida licencia por RUT y dominio en endpoint:
  - `licencias_medicas.ashx`
- Para remuneración mensual se consulta:
  - `liquidaciones_procesos.ashx` (se toma `REMUNERACIONES`)
  - `liquidaciones/detalle.ashx`
- `monto_mes = valor_dia * dias_habiles`.
- `monto_devolucion = suma(monto_mes)` por cada fila original.
- `valor_dia = remuneracion / 30`.

## Ejecutar local

```bash
cd fix_dte
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Abrir: `http://localhost:5000`

## Pruebas

```bash
cd fix_dte
python3 -m unittest discover -s tests -v
```
