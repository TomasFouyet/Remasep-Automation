# Validación Windows de la ampliación Medinet — 13 de noviembre de 2026

Rama de partida: `feature/medinet-expanded-scope`,
`be82e0d13dc9e3219afddf41a42b3284a7e53e77`, working tree limpio.
Se validó con Python 3.12.10 en `.venv-win`, pywin32 y Excel Desktop 16.0.
No se modificaron reglas, manifiesto, plantilla ni datos para igualar el histórico.

## Cálculo y escritura real

Export directo de noviembre, filtrado a septiembre: 1.821 citas ambulatorias,
595 excluidas por ESTADO, 1.226 procesadas y 0 inválidas bloqueantes.
Se producen 1.207 destinos únicos y `completeness.ok=true`.

- Las 1.122 escrituras originales conservan el digest
  `557ef5fb3171f2d9fe6c6e7832e4cd27540e542e8e051b80160fc8e00333117c`.
- Digest de las 1.207 escrituras:
  `9dbdae29b64de3c839b97a18c1c1b52eeda1ed1bde3afdf859952b4e800e8b32`.
- Las 85 reglas adicionales se evaluaron también con Excel real en un libro
  temporal no guardado: 85 coincidencias. Las 1.226 edades coinciden con Excel.
- B2 ANEXO: `14,1,6,15,1,4,135,55`; rehabilitación Santiago `81+54=135`,
  Telemedicina `20+35=55`.
- Writer productivo desde plantilla oficial en blanco: nuevo
  `outputs/REMASEP_2026_09_EXTENDED_DRAFT.xlsm`.
- Reapertura COM: 1.207 matched, 0 mismatch, 0 missing.
- Fingerprint `stf:dc624775927d4d4d`; 11.466 fórmulas y contrato semántico
  preservados, sin adiciones, eliminaciones ni cambios de texto.
- `WRITER_INTEGRITY_PASS`, `VBA_INTEGRITY_PASS`, `CLEANUP_OK`.
- SHA256 de salida:
  `31b0c38d3623fac8056dcca6e9a88027b48091d39793f8ff45f5f16d3705c019`.

## Comparación e incompletitud funcional

Sobre los 1.207 destinos: 999 exactos, 34 cero/blanco equivalentes, 174
diferencias reales, 0 referencias no disponibles (85,5841% equivalencia).
Las 85 adiciones tienen 56 exactos y 29 diferencias por edad.

La sustitución diagnóstica de AF por su caché legacy reproduce 173 diferencias.
Médico menor de 10: 55 recalculado frente a 81 manual; 26 atenciones de edad
recalculada >=10 tenían cero en caché. Enfermería: 22 frente a 47, con 25 casos
análogos. Los datos individuales permanecen en memoria y no se publican.
`REMASEP_OD!AF89` conserva la discrepancia de datos/criterio legacy documentada
en la auditoría anterior, no atribuible únicamente a AF.

También se compararon 21.630 coordenadas pobladas en las 11 hojas, separando
destinos, entradas manuales fuera de alcance, fórmulas y contenido estático.
Las expresiones de las fórmulas coinciden con las del histórico. Los resultados
derivados pueden diferir por edades, entradas manuales, metadatos y la consulta
WEBSERVICE de la portada; no se recalculó ni guardó el histórico.

CONTROL tiene **49** alertas: NOMBRE 4, REMASEP 01 17, REMASEP_OD 28.
A las 38 alertas complementarias anteriores se añaden siete: N15/O15/N16/O16
(pueblos originarios/migrantes) y las validaciones de W/X en filas 55, 75 y 89
(consultas resueltas no presenciales, por teléfono u otras plataformas).
Estas siete validaciones involucran diez entradas vacías fuera del manifiesto.
La modalidad o sucursal por sí solas no demuestran la condición clínica de
consulta resuelta ni ese desglose; no se rellenaron ceros para eludir CONTROL.

El export no aporta campos explícitos de embarazo, identidad indígena o
migración. Tampoco se demostró una regla de cirugía mayor ambulatoria o cirugía
odontopediátrica para `REMASEP_OD!Q61` y `S81`. No se incorporan esas celdas.
**El libro sigue siendo DRAFT / NOT_FOR_SUBMISSION. No apto para envío.**

## Correcciones demostradas y pruebas

La primera suite Windows terminó con 985 aprobadas, 7 fallidas y 4 omitidas.
Se corrigieron:

1. Checker de distribución: inspeccionaba rutas absolutas y confundía el
   directorio padre `AppData/Local` con una fuga `data/local`. Ahora usa rutas
   relativas y componentes; la fuga real sigue rechazándose. Dos regresiones
   cubren ambos casos y el nombre AppData dentro del bundle.
2. Resolución de referencias: la etiqueta de nombre exacto dependía de la
   semántica case-insensitive de Windows. Ahora se compara el nombre literal
   de la entrada antes de casefold y se conserva el rechazo de ambigüedades.
3. Test UI sin Excel: se inyecta explícitamente esa capacidad y se aíslan las
   rutas de diagnóstico. No se cambia el preflight productivo.
4. Dos pruebas históricas mezclaban los 1.122 destinos originales con el
   contrato de 1.207. Se conserva íntegro el golden de julio (incluidos 187
   valores no cero y suma 928) en su cohorte original, y se compara el contrato
   ampliado con su auditoría completa. No se sustituyeron goldens por resultados
   observados para forzar aprobación.
5. El subproceso `python -I` ignoraba PYTHONUTF8 y emitía cp1252; el padre UTF-8
   perdía stdout con UnicodeDecodeError en un hilo lector. Se fija `-X utf8`
   y la decodificación explícita, y se exige observar `SMOKE OK`.

Suite final Windows: **995 aprobadas, 0 fallidas y 3 omitidas**, sin warnings
pytest. Se activó `REMASEP_VALIDATE_AF_EXCEL_COM=1` y se convirtió
`PytestUnhandledThreadExceptionWarning` en error. Julio pasó; abril–junio no
están disponibles. Las otras omisiones son el fixture Workbook_Open ausente
y chmod/ACL en Windows.

Persisten mensajes nativos `0x80010108` al cerrar COM durante pytest, sin aborto
del proceso ni fallo de las comprobaciones de integridad. No se silenciaron.
El aborto previo `0xC0000005` no se reprodujo en estas suites; no se declara
corregida una causa que no se aisló. No se tocaron procesos Excel del usuario.

Las evidencias locales, comandos, hashes de originales, tablas y logs sin PII
están en `artifacts/qa_medinet_extended_20261113/`. Los resultados del build,
arranque y eventual generación del EXE se registran allí por separado, después
de este gate de generación e integridad; la validación desde Python no los
sustituye. No se hizo merge.
