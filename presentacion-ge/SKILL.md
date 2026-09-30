---
name: presentacion-ge
description: Genera presentaciones PowerPoint (.pptx) de Grupo Educativo sobre la plantilla oficial (Plantilla GE v0.6), a partir de un guion JSON. Usar cuando alguien de GE pida una PPT, presentación o deck institucional en PowerPoint, para clientes, licitaciones, reuniones o capacitaciones.
---

# Presentaciones GE (PowerPoint sobre la plantilla oficial)

Esta skill produce un `.pptx` que **hereda la plantilla oficial** (`recursos/Plantilla GE v0.6.potx`): layouts, logo, fondo, Red Hat Display incrustada y paleta del Brandbook 2026. Claude no diseña la lámina desde cero: escribe un **guion JSON** y el script `scripts/generar_ppt.py` lo convierte en láminas usando los layouts reales de la plantilla.

Si la persona pide el deck en **Claude Slides** (o no nombra formato y el artifact Slides está disponible y no necesita .pptx), usar el artifact Slides con la identidad GE descrita en "Identidad GE" más abajo. Esta skill es para cuando se necesita el archivo PowerPoint.

## Flujo

1. **Entender el encargo** (si falta algo, preguntar una vez): audiencia, propósito, duración o n.º de láminas, material fuente (bases técnicas, informe, apuntes).
2. **Leer el material fuente** y extraer lo sustantivo. Nunca inventar cifras, citas ni nombres: lo que falte va como `[__]` o `[nombre]` y se lista al entregar.
3. **Proponer el guion** en el chat, en una tabla breve: n.º, tipo de lámina, título, idea central. Esperar visto bueno si la persona está presente; si no, seguir y declarar los supuestos.
4. **Escribir el JSON** (esquema abajo) y generar:
   ```bash
   python3 scripts/generar_ppt.py guion.json "Nombre presentación.pptx"
   ```
   (el script usa la `.potx` de `recursos/`; para otra versión, `--plantilla RUTA`).
5. **Leer los AVISOS DE DENSIDAD** que imprime el script. Corregir el guion (recortar, dividir la lámina, cambiar a tarjetas) y regenerar hasta que no queden avisos.
6. **Verificar visualmente**: convertir a PDF (`soffice --headless --convert-to pdf`) y revisar las imágenes de las láminas (solapamientos, texto cortado, láminas vacías). Corregir y regenerar.
7. Entregar el `.pptx` + lista de placeholders `[__]` que la persona debe completar.

## Esquema del guion

```json
{
  "titulo": "Nombre del deck",
  "autor": "Grupo Educativo",
  "laminas": [ { "tipo": "...", "...": "...", "notas": "notas del orador (opcional)" } ]
}
```

En cualquier texto: `**negrita**` y `==destacado==` (se pinta magenta; usar máximo una vez por lámina).

| tipo | Layout de la plantilla | Campos |
|---|---|---|
| `portada` | Título imágenes incrustadas (fotos GE) | `titulo` (≤ 60 car.), `subtitulo`, `variante`: `"fotos"` o `"teal"` + `imagenes`: [3 rutas] para fotos propias |
| `seccion` | Sección (teal) / Sección (blanco) | `titulo`, `subtitulo`, `fondo`: `"teal"` (def.) o `"blanco"` |
| `agenda` | Sólo título | `titulo`, `items`: [{`texto`, `min`}] (3–7) |
| `contenido` | Un elemento (blanco/teal) | `titulo`, `intro` (opc.), `puntos`: ["a", ["sub a1"], "b"] (≤ 6), `fondo`, `tamano` (def. 20) |
| `dos_columnas` | Dos elementos (blanco) | `titulo`, `izquierda`/`derecha`: {`subtitulo`, `color`, `puntos`} |
| `tarjetas` | Sólo título | `titulo`, `intro` (opc.), `numerar`: bool, `tarjetas`: [{`titulo`, `texto`, `color`}] (2–8) |
| `cifras` | Sólo título | `titulo`, `cifras`: [{`valor`, `etiqueta`, `detalle`, `color`}] (2–4), `nota` |
| `proceso` | Sólo título | `titulo`, `pasos`: [{`titulo`, `texto`, `cuando`}] (3–5) |
| `tabla` | Sólo título | `titulo`, `encabezados`, `filas`, `anchos` (proporciones), `fuente` (≤ 8 filas) |
| `cita` | Sólo título | `titulo` (opc.), `texto`, `fuente` (APA) |
| `pregunta` | Vacía sin logo | `texto` (usar `==x==`), `bajada` (opc.) — lámina de impacto |
| `imagen` | Sólo título | `titulo`, `imagen` (ruta local), `pie` |
| `cierre` | Cierre (teal/blanco) | `fondo`, `texto` (opc., contacto o llamado) |

Colores válidos para `color`: `teal`, `magenta`, `verde`, `azul`, `violeta`, `slate`. Si no se indica, rotan teal → magenta → verde → azul. El verde se usa automáticamente en su sombra `#8B9000` cuando es texto.

## Reglas editoriales (obligatorias)

- **Estructura típica**: portada → agenda/objetivos → (sección teal → 2–5 láminas de contenido) × n → cierre. Las secciones teal marcan los grandes bloques; la sección blanca, subbloques.
- **Una idea por lámina.** Si no cabe, se divide; nunca se achica la letra por debajo de los defaults.
- **Máximo 6 viñetas**; si las ideas son paralelas, usar `tarjetas`; si son números, `cifras`; si son etapas, `proceso`; si comparan atributos, `tabla`.
- **Títulos cortos** (≤ 45 caracteres, 1 línea ideal) y con la misma gramática en todo el deck (todos sustantivos o todos preguntas).
- **Ritmo visual**: no más de 3 láminas seguidas del mismo tipo; intercalar una `pregunta` o `cifras` cada tanto.
- **Fuentes**: toda cifra o cita textual lleva fuente en formato APA (en `fuente`, `nota` o `pie`).
- **Español de Chile**, lenguaje institucional claro; sin anglicismos innecesarios.
- No editar la plantilla ni sus layouts; no agregar logos, fondos ni tipografías externas.

## Identidad GE (fuente: Design System «Grupo Educativo» en Claude Design, default de la organización)

- **Tipografía**: Red Hat Display. Títulos SemiBold 600; párrafos Regular 400; texto largo Light 300; citas en itálica; etiquetas Bold 700 en mayúsculas pequeñas con letter-spacing.
- **Primarios**: Teal `#005062` · Magenta `#C0006E` · Verde `#BBC100` · Slate `#414E55`.
- **Secundarios** (acentos por categoría): Azul `#30486B` · Violeta `#60306C` · Púrpura `#90186D` · Frambuesa `#AB2F52` · Café `#965F37` · Oliva `#818F1B` · Verde claro `#90B86D` · Verde vivo `#51A71A` · Verde medio `#368F35` · Verde bosque `#18774F`.
- **Neutros**: Blanco `#FFFFFF` · Gris fondo `#F5F6F7` · Gris borde `#E1E6E9` · Gris texto `#6B7A82`.
- **Roles**: teal = títulos y superficies corporativas; slate = texto; magenta = acentos y énfasis; verde = solo detalles (nunca texto sobre blanco ni fondo de números). Rotación de acentos por defecto: teal → magenta → azul → violeta.
- **Estética**: fondos blancos o gris fondo; tarjetas con borde gris y **barra izquierda** del color de acento; **esquinas rectas** (sin redondeos); viñetas cuadradas magenta; mucho espacio en blanco.
- **Líneas de acción** (icono + acento): 01 Infancias → magenta · 02 Formación Técnico Profesional → verde vivo · 03 Formación continua → café · 04 Mejoramiento educativo → violeta. En el guion, `"linea": "infancias" | "tp" | "formacion continua" | "mejoramiento"` en una tarjeta, cifra o paso aplica su color.
- **Logo**: versión color sobre fondo claro, blanca sobre teal; nunca deformar, recolorear ni sombrear.
- En el .pptx los títulos de lámina los define la plantilla v0.6 (Slate bold): no sobrescribirlos.

## Variante Claude Slides (mismo sistema, en el navegador)

Cuando se pida en Claude Slides, replicar el template «Deck de presentación» del Design System en el lienzo de 1920×1080:

- Subir como assets del deck los logos (`recursos/ge-logo-*.png`) y, para la portada, las tres fotos de la .potx (`ppt/media/image2–4`).
- Fuente: Red Hat Display (Google Fonts). Fondo `#F5F6F7`; secciones de bloque y cierre en teal `#005062`.
- **Contenido**: título `left:132px; top:80px; width:1452px`, 54px/600 teal; logo color `left:1625px; top:57px`, alto 107px; cuerpo desde `top:288px`, 32px slate con viñetas ■ magenta; pie a 24px gris texto con fecha · «Grupo Educativo · Soluciones integrales» · n.º.
- **Sección teal/blanca**: logo arriba a la izquierda (`132,96`, alto 107px); título 80px/600 en `top:610px` (blanco sobre teal, slate sobre claro); bajada 34px/300 en `top:812px`.
- **Portada**: fotos en `top:81px`, alto 492px (x 3/513/1449, anchos 471/894/471); título 80px/600 teal en `top:610px`; subtítulo 34px/300; logo abajo a la derecha (`1560,924`).
- **Cierre**: fondo teal, logo blanco grande centrado (`348,249`, ancho 1224px).
- Lámina de contenido teal: solo para 1–2 mensajes clave por presentación, viñetas ■ verde sobre teal.
- Mismos tipos de lámina y reglas editoriales de esta skill; texto de cuerpo ≥ 28px.

## Recursos y ejemplo

- `recursos/Plantilla GE v0.6.potx` — plantilla oficial (fuente de layouts, logo, fondos y fuentes). Al publicar una nueva versión, reemplazar este archivo y revisar que los nombres de layout del diccionario `LAYOUT` en el script sigan existiendo.
- `recursos/ge-logo-color.png`, `ge-logo-blanco.png`, `ge-logo-blanco-grande.png` — logos para la variante Claude Slides.
- `ejemplos/presentaciones-ge-con-claude.json` + `ejemplos/Presentaciones GE con Claude.pptx` — guion de referencia que usa casi todos los tipos de lámina, y su resultado. Úsalo como modelo de estructura y tono antes de escribir un guion nuevo.

## Problemas conocidos

- El render de LibreOffice puede diferir levemente de PowerPoint; si algo queda justo, preferir recortar texto.
- El estimador de densidad es conservador: un aviso "justo" (diferencia < 0.2") suele caber, pero conviene revisar la imagen.
