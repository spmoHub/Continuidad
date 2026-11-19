CONTINUIDAD 2027 · V11.6
CORRECCIÓN DE VALIDACIÓN + PESTAÑA INDEPENDIENTE

PRONÓSTICO OPERATIVO
NO CAMBIA.

Sigue siendo:
Pronóstico total =
Pronóstico sobre BASE_VIVA.xlsx actual
+
respondientes que hacen match con INSCRITOS_2027.xlsx

CORRECCIÓN DE VALIDACIÓN

La versión anterior limitaba los inscritos de la validación a quienes también
estaban en BASE_VIVA_INICIAL.xlsx. Eso excluía inscritos reales que sí habían
respondido la encuesta.

La lógica correcta de validación ahora es:

UNIVERSO DEL MODELO =
todos los respondientes únicos de la hoja Alumnos

POSITIVOS OBSERVADOS =
respondientes que hacen match con INSCRITOS_2027.xlsx

BASE_VIVA_INICIAL.xlsx =
solo una fotografía adicional del funnel; NO filtra los inscritos del análisis.

Por lo tanto, si existen:
- 763 inscritos que hacen match con respondientes
- 412 de ellos estaban en Base Viva inicial
- 351 no estaban en Base Viva inicial

la validación usa LOS 763 COMPLETOS.

NUEVA PESTAÑA
Se agregaron dos pestañas:
1. Pronóstico y seguimiento
2. Validación del modelo

La pestaña de validación muestra:
- Respondientes del modelo
- Ya inscritos
- Conversión observada parcial
- Inscritos que estaban en Base Viva inicial
- Inscritos fuera de Base Viva inicial
- % de inscritos que eran Alta/Muy alta
- Respondientes en Base Viva inicial
- Siguen en Base Viva actual
- Fuera de Base Viva y no inscritos
- Probabilidad promedio de todos los respondientes
- Probabilidad promedio de los ya inscritos
- Tabla de conversión por propensión

TABLA POR PROPENSIÓN
Para cada categoría:
- Respondientes
- Inscritos
- Conversión observada
- En Base Viva inicial
- Inscritos fuera de Base Viva inicial
- Probabilidad promedio del modelo

IMPORTANTE
La validación sigue siendo PARCIAL:
Inscrito = positivo observado.
No inscrito todavía ≠ negativo final.

Por eso no se muestra todavía Accuracy/Precision/Recall/F1 final.

ARCHIVOS A REEMPLAZAR EN GITHUB
- index.html
- styles.css
- app.js
- .nojekyll

ARCHIVOS DE DATOS
- base_continuidad_2027_dashboard.xlsx
- BASE_VIVA_INICIAL.xlsx
- BASE_VIVA.xlsx
- INSCRITOS_2027.xlsx
