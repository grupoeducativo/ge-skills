# ge-skills

Skills de Claude Code para el equipo de Grupo Educativo.

## Instalación

Agrega esto a tu `~/.claude/settings.json`:

```json
"extraKnownMarketplaces": {
    "ge-skills": {
        "source": {
            "source": "github",
            "repo": "grupoeducativo/ge-skills"
        }
    }
}
```

Luego instala las skills desde Claude Code con `/install-skill`.

## Configuración del token Moodle

Cada integrante del equipo debe generar su propio token en:  
**aprendizajeservicio.cl → Administración → Servidor → Servicios web → Gestionar tokens**

Guárdalo como variable de entorno en tu máquina:

```bash
# En ~/.zshrc o ~/.bashrc
export MOODLE_TOKEN="tu_token_aqui"
```

## Skills disponibles

| Skill | Descripción |
|---|---|
| `moodle-reportes` | Genera reportes y visualizaciones de avance desde Moodle |
