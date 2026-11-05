# REVISION-TOPES-HORARIO

Horario interactivo de Ingeniería UANDES para revisar topes entre NRC.
Piloto para el periodo **202710**.

- Una sola vista, sin login: cualquiera con el enlace puede entrar.
- Filtros por área, materia, curso, sección, título, NRC, LC, sala, tipo y profesor.
- **Cargar ramos inscritos por RUT de alumno**: filtra automáticamente los NRC
  del alumno (reporte UG307) y revisa sus topes.
- Revisión de topes de la selección actual, con lista de topes y celdas marcadas.
- Descarga del Excel de la hoja HORARIO.

## Estructura

```
index.html, app.js, styles.css   Sitio estático (GitHub Pages)
data/config.json                 Periodo y URL del Web App (lo escribe Apps Script)
data/data.json                   Filas de la hoja HORARIO (lo escribe Apps Script)
apps-script/Web.gs               Parámetros, UG307, Web App, publicación, descarga
apps-script/Horario.gs           Lógica de construcción del horario (ver paso 2)
```

> El `data/data.json` incluido es una **muestra con datos 202620** para probar
> el sitio. Se reemplaza en la primera publicación desde Apps Script.

## Puesta en marcha

### 1. Repositorio y GitHub Pages
1. Crear el repositorio `REVISION-TOPES-HORARIO` y subir estos archivos.
2. Settings → Pages → Deploy from branch → `main` / root.
3. Crear un token **fine-grained** con permiso *Contents: Read and write*
   solo sobre este repositorio.

### 2. Google Sheet y Apps Script
1. Crear el Google Sheet del periodo 202710 (hojas `NRC POR PERIODO`, etc.).
2. Extensiones → Apps Script. Crear dos archivos:
   - `Web.gs`: pegar `apps-script/Web.gs` completo.
   - `Horario.gs`: pegar del script anterior el bloque que va **desde
     `function normalizeName(str) {` hasta el final de `updateSalasPruebas()`**
     (justo antes del comentario `// REVISAR AJUSTES DPSA`). Ese bloque no cambia.
     No copiar nada de lo que está antes de `normalizeName`: ya está en `Web.gs`.
3. Recargar el Sheet. Aparece el menú **Macros**.
4. Macros → Configurar MAESTRO fuente · Configurar GitHub.
5. Completar `PARAMETROS_HORARIO` (PERIODO, fechas de CLAS y AYUD/LAB).

### 3. Web App
1. Implementar → Nueva implementación → Aplicación web.
   *Ejecutar como:* Yo · *Quién tiene acceso:* Cualquier persona.
2. Macros → Configurar URL Web App (pegar la URL `/exec`).

Al cambiar `Web.gs` hay que crear una **nueva versión** de la misma
implementación (Gestionar implementaciones → editar → Nueva versión) para que la
URL no cambie.

### 4. Inscripciones (UG307)
1. Archivo → Importar → subir el Excel UG307 → *Insertar hojas nuevas*.
2. Renombrar la hoja importada a `UG307`.
3. Macros → Importar inscripciones UG307.
   - Se crea la hoja `INSCRIPCIONES` con solo RUT, NRC, materia, curso y asignatura
     del periodo configurado (sin nombres).
   - Se ofrece eliminar la hoja `UG307` original.

Repetir cada vez que haya un UG307 nuevo (se invalida el caché automáticamente).

### 5. Horario y publicación
Macros → Crear Horario → Llenar Salas Horario → Actualizar Salas Pruebas →
**Publicar/Actualizar web**. Antes, *Diagnosticar* muestra si falta algo.

## Privacidad

- La relación RUT → NRC **no se publica** en GitHub. La web consulta un RUT a la
  vez al Web App, que responde solo NRC, materia, curso y asignatura.
- Límite de consultas por minuto y caché de 10 minutos por RUT.
- El token de GitHub vive solo en Script Properties. No crear hojas con tokens.

## Reglas de topes

Se comparan todos los pares de filas de la selección, salvo: mismo NRC, mismo
curso (otra sección), misma LC o cursos homologables (misma materia y título,
distinto número de curso).

| Tipo | Condición |
|---|---|
| Clase / clase | CLAS, AYUD o LAB del mismo día, con horas que se cruzan **y** periodos INICIO–FIN que se cruzan |
| Evaluación / evaluación | PRBA o EXAM en la misma fecha y con horas que se cruzan |
| Evaluación / clase | PRBA/EXAM que cae sobre una clase de otro curso, solo si la evaluación está fuera del horario regular de su propio curso (si está dentro, el tope ya aparece como Clase / clase) |

Los NRC inscritos que no están en la hoja HORARIO (por ejemplo, CEG u otros
cursos de otras facultades) se listan aparte y no se revisan.

## Nota sobre fechas de primer semestre

`createHorario` resta 2 días al inicio de CLAS/AYUD/LAB en cursos con
`Plan Común = 1`. En la versión anterior ese ajuste no se aplicaba, porque las
fechas de `PARAMETROS_HORARIO` se leían mal; en `Web.gs` esa lectura está
corregida, así que **ahora sí se aplicará**. Si no corresponde, cambiar en
`Horario.gs` la línea `shiftDateString_(ini, -2, tz)` por `0`.
