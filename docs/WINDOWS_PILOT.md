# REMASEP Automation — Guía del piloto

**Versión 0.1.0** · requiere **Windows** y **Microsoft Excel Desktop**
instalados en el computador.

Esta guía no requiere conocimientos técnicos. No hace falta instalar Python,
ni usar una terminal, ni nada parecido: sólo copiar una carpeta y hacer doble
clic.

---

## 1. Instalar

1. Copia la carpeta **`REMASEP`** completa (la que contiene `REMASEP.exe`) a
   una ubicación de tu computador, por ejemplo el Escritorio.
2. Eso es todo — no hay instalador que ejecutar.

## 2. Abrir la aplicación

Haz doble clic en **`REMASEP.exe`** dentro de esa carpeta.

## 3. Generar el informe del mes

1. En la pantalla de inicio, pulsa **Nuevo informe**.
2. Pulsa **Seleccionar archivo** junto a "Archivo de Medinet" y elige el
   export "Detalle de citas" del mes que quieres procesar.
3. La aplicación detecta automáticamente el período (mes/año) a partir de las
   fechas del archivo. Si el archivo tiene varios meses, te pedirá confirmar
   cuál procesar.
4. Pulsa **Seleccionar archivo** junto a "Plantilla REMASEP" y elige la
   plantilla oficial (`.xlsm`).
5. Pulsa **Analizar datos**.
6. Cuando termine, verás el **Resumen mensual**: cuántas citas hay en el
   período, cuántas se consideran para el REMASEP y cuántas quedan excluidas,
   con gráficos.
7. Si quieres, pulsa **Exportar resumen PDF** y elige dónde guardarlo (por
   ejemplo, el Escritorio).
8. Pulsa **Generar REMASEP**.
9. Se abrirá un diálogo **"Guardar informe REMASEP"**: elige la carpeta y el
   nombre del archivo (se sugiere uno automáticamente) y confirma.
10. Espera mientras se genera — Microsoft Excel realiza el cálculo. No
    cierres la aplicación durante este paso.
11. Al terminar, verás el resultado con el nombre del archivo, el período y
    la verificación de integridad. Puedes pulsar **Abrir Excel** para verlo
    directamente o **Abrir carpeta** para ir a donde se guardó.

**Importante:** Microsoft Excel Desktop es necesario para generar el archivo
REMASEP. Si no está instalado, la aplicación lo indicará con un mensaje claro
en el paso de generación (los pasos anteriores — cargar Medinet, analizar,
ver el resumen, exportar el PDF — funcionan igual sin Excel).

Si ya existe un archivo con el nombre que elegiste, la aplicación te lo dirá
y te pedirá elegir otro nombre — nunca sobrescribe un archivo existente.

## 4. Cerrar y volver a abrir

Puedes cerrar la aplicación en cualquier momento y volver a abrir
`REMASEP.exe` cuando quieras. No queda nada "a medio hacer": cada informe se
genera de principio a fin en una sola sesión.

## 5. Si algo sale mal

Los mensajes de la aplicación están escritos para ser comprensibles (por
ejemplo: *"No pudimos leer el archivo de Medinet"*, *"La plantilla no es
compatible"*). Si ves un mensaje así, sigue la indicación que ofrece (elegir
otro archivo, elegir otra plantilla, etc.).

Si necesitas soporte técnico, comparte con quien te asista **la versión**
(0.1.0, indicada arriba) — no hace falta enviar ningún archivo con datos de
pacientes.

---

*Esta guía cubre el flujo normal de uso. La documentación de cómo se
construyó el ejecutable — para quien mantiene la aplicación, no para quien la
usa — está en [`WINDOWS_PACKAGING.md`](WINDOWS_PACKAGING.md).*
