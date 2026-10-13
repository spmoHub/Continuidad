TABLERO CONTINUIDAD 2027 · V7
==============================

CAMBIOS DE ESTA VERSIÓN
-----------------------
- Semáforo de NPS, Orgullo, Seguridad e Infraestructura:
  * < 7  = rojo
  * 7 a 8.9 = amarillo
  * >= 9 = verde
  Esto interpreta la indicación de que 7 y 8 deben quedar en amarillo.
- Se agregó un filtro superior para encontrar rápidamente oportunidades de convencimiento.
- El orden predeterminado del explorador es "Mayor oportunidad primero".
- Se agregó una guía al final del tablero que explica Propensión relativa, Potencial estructural,
  Brecha de oportunidad y Prioridad de convencimiento.
- Egresados esperados 2027 se retiró temporalmente del tablero y del cálculo absoluto.
- Mientras se valida ese dato, el volumen de continuadores se proyecta sobre la Base Viva actual.
- Programa ofertado usa únicamente `programa_ofertado` (1 = Sí, 0 = No).

ARCHIVOS DE DATOS
-----------------
El tablero lee automáticamente, sin pedir carga manual:
- base_continuidad_2027_dashboard.xlsx
- BASE_VIVA.xlsx

Si actualizas alguno, conserva exactamente esos nombres.

ABRIR LOCALMENTE
----------------
1. Descomprime la carpeta completa.
2. Haz doble clic en `abrir_tablero.bat`.
3. El navegador se abrirá automáticamente.
4. El botón "Actualizar desde Excel" vuelve a leer ambos archivos.

PUBLICAR EN GITHUB PAGES
------------------------
Esta versión también puede funcionar como sitio estático en GitHub Pages.
Sube al MISMO NIVEL del repositorio al menos estos archivos:
- index.html
- styles.css
- app.js
- base_continuidad_2027_dashboard.xlsx
- BASE_VIVA.xlsx

Después configura Pages para publicar desde la raíz de la rama elegida.
El navegador leerá directamente los dos Excel mediante JavaScript.

IMPORTANTE: `index.html` por sí solo NO es suficiente. También necesita styles.css, app.js y los dos Excel.
La librería para leer XLSX se carga desde jsDelivr, por lo que el navegador necesita conexión a internet.

PRIVACIDAD / GITHUB PAGES
-------------------------
Este tablero contiene consulta individual por matrícula y los archivos Excel quedan alojados junto al sitio.
Si el GitHub Pages es público, cualquier persona con acceso al sitio podría descargar directamente los Excel,
aunque el repositorio sea privado. Por eso NO se recomienda publicar esta versión en Pages público.
Úsala únicamente en un sitio con autenticación/control de acceso (por ejemplo Pages privado compatible,
Azure Static Web Apps con Entra ID, u otro hosting interno protegido).

PRONÓSTICO OPERATIVO ACTUAL
---------------------------
- El modelo y la propensión parten de quienes respondieron la encuesta.
- El volumen 2027 se aplica temporalmente a la Base Viva actual de cada campus.
- La tasa de referencia usa el punto medio entre el promedio histórico ponderado 2024-2026
  y el límite superior del modelo calibrado.
- En campus sin histórico comparable se usa la señal calibrada del modelo.
- El resultado principal se muestra en número absoluto de continuadores esperados.
