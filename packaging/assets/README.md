# Icono — pendiente

No existe todavía un icono aprobado de "REMASEP Automation" / Fundación
Gantz. Por eso este piloto **no** incluye un `.ico` — no se generó ni
descargó branding arbitrario (Sprint 3.11, Fase 7).

`packaging/remasep.spec` ya busca `packaging/assets/remasep.ico`
automáticamente (usa el icono por defecto de PyInstaller si el archivo no
existe). Cuando exista un icono aprobado:

1. Colocar el archivo en `packaging/assets/remasep.ico` (formato `.ico`,
   idealmente con tamaños 16/32/48/256 px).
2. Volver a construir: `.\scripts\build_windows.ps1`.

No hace falta tocar el `.spec` ni el script de build.
