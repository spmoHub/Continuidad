CONTINUIDAD 2027 · V11 · INSCRITOS + NOWCAST

ARCHIVOS DE CÓDIGO QUE DEBES SUBIR/REEMPLAZAR EN GITHUB:
- index.html
- styles.css
- app.js
- .nojekyll
- base_continuidad_2027_dashboard.xlsx (se incluye sin cambios; puedes conservar el que ya tienes si es el mismo)

ARCHIVOS DE DATOS QUE DEBEN EXISTIR EN LA MISMA CARPETA:
- BASE_VIVA.xlsx
- INSCRITOS_2027.xlsx

IMPORTANTE:
1) NO necesitas borrar columnas de INSCRITOS_2027.xlsx.
2) Para esta versión se usan principalmente:
   - Número de matrícula
   - Campus
3) La hoja recomendada se llama: Inscritos_2027
   Aun así, si tiene otro nombre, el código intentará usar la primera hoja.
4) INSCRITOS_2027.xlsx debe ser acumulativo: todos los inscritos confirmados a Ago-27 hasta la fecha.
5) El tablero deduplica por matrícula.
6) Si una matrícula aparece en BASE_VIVA.xlsx y en INSCRITOS_2027.xlsx:
   - cuenta como inscrita confirmada;
   - se excluye del pronóstico probabilístico;
   - NO se cuenta dos veces.

NUEVA LÓGICA:
Pronóstico total Ago-27
= inscritos confirmados
+ continuadores esperados adicionales entre la Base Viva pendiente.

KPI:
- Base Viva actual
- Ya inscritos
- Respondentes pendientes
- Pronóstico total Ago-27
- Oportunidades de convencimiento

El botón "Actualizar desde Excel" vuelve a leer:
- base_continuidad_2027_dashboard.xlsx
- BASE_VIVA.xlsx
- INSCRITOS_2027.xlsx

La lista/ficha individual sigue mostrando alumnos con encuesta/modelo que continúan pendientes en Base Viva.
