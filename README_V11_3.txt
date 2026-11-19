CONTINUIDAD 2027 · V11.3
TODOS LOS RESPONDIENTES + ESTATUS BASE VIVA/INSCRITO + FILTRO DE PROPENSIÓN

FUENTES
- base_continuidad_2027_dashboard.xlsx
- BASE_VIVA.xlsx
- INSCRITOS_2027.xlsx

REGLAS

1) EXPLORADOR DE ALUMNOS
Ya NO se eliminan los respondientes que no hagan match con Base Viva.

Se conservan todos los respondientes únicos de la hoja Alumnos.

Cada matrícula tiene dos banderas independientes:
- En Base Viva: Sí / No
- Inscrito: Sí / No

Esto permite distinguir:
- Sigue activo en funnel.
- Ya se inscribió.
- Ya no está en Base Viva y tampoco aparece inscrito.
- Caso de auditoría: aparece simultáneamente en Base Viva e Inscritos.

2) PRONÓSTICO OPERATIVO
No cambia:
Pronóstico total = pronóstico sobre Base Viva + respondientes que hacen match con INSCRITOS_2027.

Base Viva no se modifica ni se le restan inscritos, porque se asume que ya viene depurada.

3) TENDENCIA 2027
El punto 2027 de la gráfica usa:
Pronóstico total / total de respondientes del dashboard.

Por campus:
Pronóstico total del campus / respondientes del campus.

4) PROPENSIÓN
La gráfica de propensión usa TODOS los respondientes del campus seleccionado.

Las barras son clicables y también se muestran botones:
- Todas
- Muy baja
- Baja
- Media
- Alta
- Muy alta

Al seleccionar una categoría, se filtra el explorador de alumnos.

5) FILTROS OPERATIVOS
Se agregan:
- Base Viva: Todos / En Base Viva / No en Base Viva
- Inscripción: Todos / Inscrito / No inscrito

Los filtros se pueden combinar con propensión, campus, matrícula y oportunidad.

6) VER OPORTUNIDADES
El botón "Ver oportunidades" se limita a:
- En Base Viva = Sí
- Inscrito = No
- Prioridad de convencimiento Media-Alta / Alta / Muy alta

Así no manda al asesor alumnos que ya se inscribieron o que ya no están en el funnel.

ARCHIVOS DE CÓDIGO A REEMPLAZAR EN GITHUB
- index.html
- styles.css
- app.js
- .nojekyll

Mantén tus Excel actuales en la misma carpeta.
