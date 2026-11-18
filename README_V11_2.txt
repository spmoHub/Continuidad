CONTINUIDAD 2027 · V11.2
REGLA CORRECTA: BASE VIVA + INSCRITOS DE RESPONDIENTES

ARCHIVOS DE DATOS QUE DEBEN ESTAR EN GITHUB:
- base_continuidad_2027_dashboard.xlsx
- BASE_VIVA.xlsx
- INSCRITOS_2027.xlsx

INSCRITOS_2027.xlsx:
- Puede conservar TODAS sus columnas.
- Puede contener Preparatoria, Profesional u otros niveles.
- Para esta lógica se usa "Número de matrícula" para el cruce y "Campus" para el desglose.
- No se cuentan todos sus registros.

LÓGICA EXACTA

1. RESPONDIENTES EN BASE VIVA
   hoja Alumnos ∩ BASE_VIVA.xlsx por matrícula

2. PRONÓSTICO BASE VIVA
   Se conserva la metodología que ya tenía el dashboard:
   - la señal del modelo proviene de los respondientes que hacen match con Base Viva;
   - el volumen se proyecta sobre la Base Viva actual.

3. INSCRITOS CONFIRMADOS DE ESTA COHORTE
   hoja Alumnos ∩ INSCRITOS_2027.xlsx por matrícula

   Solo estos matches se cuentan.
   Los inscritos de Preparatoria u otros universos que no estén en Alumnos NO se cuentan.

4. PRONÓSTICO TOTAL
   pronóstico Base Viva
   +
   inscritos confirmados con match en respondientes

IMPORTANTE
- NO se restan inscritos de BASE_VIVA.xlsx.
- Se asume que BASE_VIVA.xlsx ya quitó a quienes se inscribieron o salieron del funnel.
- El código muestra un control de auditoría si alguna matrícula de los respondientes aparece
  simultáneamente en BASE_VIVA e INSCRITOS_2027. Idealmente ese número debe ser 0.

FICHA INDIVIDUAL
- Sigue mostrando únicamente respondientes que hacen match con la Base Viva actual.
- Es decir, los alumnos que siguen pendientes y pueden ser trabajados por el asesor.

ARCHIVOS DE CÓDIGO A REEMPLAZAR EN GITHUB
- index.html
- styles.css
- app.js
- .nojekyll

El index carga app.js?v=11.2 para evitar caché del JavaScript anterior.
