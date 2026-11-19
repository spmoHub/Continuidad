CONTINUIDAD 2027 · V11.4
PROGRAMA Y CAMPUS REAL DE INSCRIPCIÓN

Mantiene toda la lógica de la v11.3:
- todos los respondientes permanecen en el explorador;
- etiquetas En Base Viva Sí/No;
- etiquetas Inscrito Sí/No;
- filtros de propensión;
- pronóstico total = pronóstico Base Viva + inscritos con match;
- tasa 2027 de tendencia = pronóstico total / respondientes.

NUEVO EN V11.4

Para cada matrícula que haga match entre:
hoja Alumnos del dashboard ∩ INSCRITOS_2027.xlsx

se recuperan del archivo INSCRITOS_2027.xlsx:
- Programa de Interés
- Campus

En la ficha individual se muestran como:
- Programa en el que se inscribió
- Campus en el que se inscribió

Estos campos NO sustituyen los datos de encuesta:
- Carrera que quiere estudiar
- Programa de interés declarado
- Campus de interés

Así se distingue claramente:
DECLARADO EN ENCUESTA vs RESULTADO REAL DE INSCRIPCIÓN.

Para alumnos no inscritos, Programa inscrito y Campus inscrito aparecen como “—”.

ARCHIVOS DE CÓDIGO A REEMPLAZAR EN GITHUB
- index.html
- styles.css
- app.js
- .nojekyll

Mantén tus Excel actuales:
- base_continuidad_2027_dashboard.xlsx
- BASE_VIVA.xlsx
- INSCRITOS_2027.xlsx
