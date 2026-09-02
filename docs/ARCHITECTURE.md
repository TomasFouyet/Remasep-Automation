# Arquitectura

## Componentes

**Extract → Transform → Metrics → Map → Load**

### Extract
Adaptadores independientes para Medinet, egresos y recursos.

### Transform
Normalizar, validar, clasificar y agregar.

Regla de seguridad:
- 0 matches: bloquear.
- 1 match: continuar.
- >1 matches: bloquear.

### Metrics
Tabla larga semántica:

```text
formulario | categoria | especialidad | sexo | edad | modalidad | valor
```

### Map
La lógica de negocio nunca conoce `G54`.
El mapping por versión traduce la métrica semántica a una celda/rango.

### Load
En Windows se usa Excel vía `DispatchEx`.

- abrir copia;
- validar estructura;
- escribir;
- recalcular;
- revisar CONTROL;
- guardar;
- cerrar solo la instancia creada por la aplicación.
