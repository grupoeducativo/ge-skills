---
name: moodle-reportes
description: Genera reportes y visualizaciones de avance de participantes desde la plataforma Moodle de Grupo Educativo (aprendizajeservicio.cl). Úsalo cuando se pidan reportes, gráficos, tablas o análisis de avance, participación o completación de actividades en Moodle.
---

# Skill: moodle-reportes

Genera reportes de avance de participantes en cursos de Moodle usando la API REST de aprendizajeservicio.cl.

## Configuración

- **Base URL**: `https://aprendizajeservicio.cl/webservice/rest/server.php`
- **Token**: leer desde variable de entorno `MOODLE_TOKEN`
- **Formato**: siempre usar `moodlewsrestformat=json`

Para obtener el token en bash:
```bash
TOKEN=$(printenv MOODLE_TOKEN)
```

Si `MOODLE_TOKEN` no está definido, detente y pide al usuario que lo configure antes de continuar.

## Funciones disponibles en la API

### Cursos
- `core_course_get_courses` — lista todos los cursos (sin parámetros)
- `core_course_get_contents&courseid=ID` — estructura de secciones y actividades

### Participantes
- `core_enrol_get_enrolled_users&courseid=ID` — inscritos con lastaccess y firstaccess

### Completación
- `core_completion_get_activities_completion_status&courseid=ID&userid=UID` — estado de completación por usuario (state: 0=incompleto, 1=completo)

### Tareas
- `mod_assign_get_assignments&courseids[0]=ID` — tareas del curso
- `mod_assign_get_submissions&assignmentids[0]=ID` — entregas por tarea

### Foros
- `mod_forum_get_forums_by_courses&courseids[0]=ID` — foros del curso
- `mod_forum_get_forum_discussions&forumid=ID` — discusiones de un foro

### Cuestionarios
- `mod_quiz_get_quizzes_by_courses&courseids[0]=ID` — quizzes del curso
- `mod_quiz_get_user_attempts&quizid=ID&userid=UID` — intentos por usuario

## Tipos de reporte

### 1. Funnel de avance (reporte estándar)

Muestra el embudo de participación en un módulo específico. Pasos:

1. Obtener inscritos del curso → contar total y con `lastaccess > 0`
2. Obtener CMIDs del módulo objetivo desde `core_course_get_contents`
3. Llamar `core_completion_get_activities_completion_status` para cada usuario (usar Python con `urllib.request` en loop, con `time.sleep(0.05)` entre llamadas)
4. Contar usuarios con al menos 1 actividad completada en el módulo
5. Obtener foros del módulo → contar participantes únicos en discusiones
6. Generar HTML con el funnel (ver plantilla más abajo)

**Definición de "ingresó al módulo"**: tiene `state=1` en al menos 1 CMID del módulo.  
**Definición de "publicó en foro"**: aparece como autor en al menos 1 discusión del foro del módulo.

### 2. Tabla de avance por participante

Para cada inscrito, muestra: nombre, último acceso, actividades completadas (N de M), y si publicó en foros. Generar como tabla markdown o CSV.

### 3. Participantes sin actividad

Lista de inscritos con `lastaccess = 0` (nunca ingresaron). Útil para seguimiento.

## Plantilla HTML — Funnel

Genera un archivo `.html` con este diseño de funnel horizontal usando barras proporcionales al 100% del total:

```html
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Funnel — [Nombre Curso]</title>
<style>
  body { font-family: -apple-system, sans-serif; background: #f5f5f5; display: flex; justify-content: center; padding: 40px 20px; }
  .container { background: white; border-radius: 12px; padding: 40px; max-width: 640px; width: 100%; box-shadow: 0 2px 16px rgba(0,0,0,0.08); }
  h1 { font-size: 20px; color: #1a1a1a; margin-bottom: 4px; }
  .subtitle { font-size: 13px; color: #888; margin-bottom: 36px; }
  .funnel { display: flex; flex-direction: column; gap: 4px; }
  .bar-wrap { display: flex; flex-direction: column; gap: 4px; margin-bottom: 4px; }
  .bar-label { font-size: 13px; color: #555; display: flex; justify-content: space-between; }
  .bar-label .name { font-weight: 500; color: #222; }
  .bar-track { height: 36px; background: #f0f0f0; border-radius: 6px; overflow: hidden; }
  .bar-fill { height: 100%; border-radius: 6px; display: flex; align-items: center; padding-left: 12px; }
  .bar-fill .count { color: white; font-size: 15px; font-weight: 700; }
  .drop { display: flex; align-items: center; margin: 2px 0 2px 6px; }
  .drop-line { width: 2px; height: 16px; background: #e0e0e0; }
  .drop-label { font-size: 11px; color: #bbb; margin-left: 10px; }
  .drop-label .diff { color: #e07070; font-weight: 600; }
  .footer { margin-top: 24px; font-size: 11px; color: #bbb; border-top: 1px solid #f0f0f0; padding-top: 16px; }
</style>
</head>
<body>
<div class="container">
  <h1>[Nombre del Curso]</h1>
  <p class="subtitle">[Módulo] · [Fecha corte]</p>
  <div class="funnel">
    <!-- Repetir por cada paso del funnel: -->
    <div class="bar-wrap">
      <div class="bar-label"><span class="name">[Etapa]</span><span>[N%]</span></div>
      <div class="bar-track">
        <div class="bar-fill" style="width:[PCT]%; background:[COLOR];"><span class="count">[N]</span></div>
      </div>
    </div>
    <div class="drop">
      <div class="drop-line"></div>
      <span class="drop-label"><span class="diff">−[DIFF]</span> [razón de caída]</span>
    </div>
    <!-- ... -->
  </div>
  <p class="footer">Datos obtenidos vía API Moodle · Grupo Educativo · [Fecha]</p>
</div>
</body>
</html>
```

**Paleta de colores** (de mayor a menor en el funnel):
- Paso 1: `#2E7D5E`
- Paso 2: `#3A9E7A`
- Paso 3: `#52BF95`
- Paso 4: `#7ED4B0`

Guardar el HTML en el directorio de trabajo actual o donde el usuario indique. Luego abrir con `open archivo.html`.

## Notas importantes

- La API de completación es una llamada por usuario — para cursos grandes (>100 inscritos) el loop toma ~30 segundos. Informar al usuario antes de ejecutar.
- El campo `lastaccess` es un timestamp Unix. `lastaccess = 0` significa que nunca ingresó.
- Los CMIDs de un módulo se obtienen desde `core_course_get_contents` — usar el campo `id` de cada módulo dentro de la sección correspondiente.
- Nunca hardcodear el token en el código generado. Siempre leerlo desde `MOODLE_TOKEN`.
