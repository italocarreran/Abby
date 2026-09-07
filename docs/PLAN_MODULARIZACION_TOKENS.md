# Plan de modularización orientado a ahorrar contexto

Este documento deja a Claude y Codex el orden acordado para reducir código
repetido y el volumen que hay que leer, **sin cambiar los ejecutables, sus
ventanas, argumentos, entradas ni salidas**. Cada migración conserva wrappers
con los nombres históricos y se prueba antes de pasar a la siguiente pieza.

## Estado resumido

| Fase | Estado | Resultado esperado |
|---|---|---|
| 1. Configuración y traspaso | **Implementada en esta rama; pendiente revisión de Claude** | Una sola implementación de persistencia y del contrato opcional del Revisor |
| 2. Lectura y utilidades comunes | Pendiente | OOXML, normalización y búsqueda segura sin copias |
| 3. Infraestructura de comparadores | Pendiente | Compartir estado, cola UI y salida Excel sin mezclar sus motores |
| 4. División interna del Revisor | Pendiente | Separar UI, estado, lectores, verificaciones y lanzamiento |

## Fase 1 — configuración y traspaso

### Alcance implementado

1. Los doce programas que usan configuración delegan en
   `__comun__/config.py`. Se conservan `get_usuario`, `leer_config`,
   `guardar_config`, `_modificar_config`, `escribir_json` y
   `escribir_json_atomico` donde ya existían, como aliases o wrappers breves.
2. `ActualizaRemplazos.py` conserva su archivo independiente
   `__config__/reemplazos_reuc.json`; los demás conservan
   `__config__/config.json`.
3. Se agregó `__comun__/traspaso.py`. Los nueve actualizadores delegan en
   `leer_argumento()`, mantienen modo manual sin argumento y conservan los
   aliases `TRASPASO_ORIGEN`/`TRASPASO_VERSION_MAX`.
4. El Revisor usa el mismo origen y versión al producir el JSON. Las claves de
   `rutas` siguen perteneciendo a cada actualizador: el módulo común solo valida
   el sobre del contrato.
5. Se agregaron pruebas del traspaso y un caso defensivo de configuración donde
   el bloque del usuario no es un diccionario.

### Puntos que Claude debe revisar

- Confirmar que los doce `CONFIG_PATH` no cambiaron.
- Confirmar que el JSON roto jamás se sobrescribe y que una escritura conserva
  las claves de otros equipos.
- Confirmar los casos de traspaso: ausente, ruta vacía, archivo inexistente,
  JSON roto, origen ajeno, versión futura y `rutas` inválida.
- Confirmar que cada actualizador sigue arrancando sin argumento y que el
  Revisor sigue escribiendo `_traspaso_actualizador.json` con versión 1.
- La prueba real de las ventanas, Excel, Access y SQL Server sigue requiriendo
  Windows y archivos de trabajo; el cambio de esta fase es anterior a esas
  operaciones.

## Fase 2 — lectura y utilidades comunes

Aplicar **una pieza por vez**, en este orden:

1. `__comun__/excel_xml.py`: extraer primero las primitivas OOXML duplicadas
   entre el Revisor y `Actualiza_Data_Access.py` (`NS_XL`, `NS_REL`, entidades,
   ubicación de hojas, expansión y lectura de columnas). Mantener en cada
   consumidor su adaptación de dominio y el fallback actual a xlwings.
2. `__comun__/texto.py`: inventariar antes todas las variantes de
   `normalizar`; separar normalización estricta de centrales/empresas/columnas
   de cualquier variante suave usada para buscar archivos.
3. `__comun__/archivos.py`: centralizar temporales de Excel, copias de Windows,
   búsqueda por patrón y preferencia del original válido más reciente. Los
   patrones específicos de cada planilla permanecen en su script.

Cada pieza debe tener pruebas stdlib y migrar primero un consumidor. Solo tras
comparar comportamiento se migra el segundo; no hacer las tres extracciones en
un único commit.

## Fase 3 — infraestructura de comparadores

Compartir por **composición**, no mediante una clase base grande:

1. persistencia de estado e inclusión mensual;
2. cola de mensajes que garantiza que solo el hilo principal toque tkinter;
3. respaldo de Excel y preservación de hojas ajenas;
4. utilidades comunes para meses, etapas y rutas.

Debe permanecer separado todo lo que define el dominio: lectura Access en
`Comparador_Etapas.py`, lectura de consolidados en `Comparador_Tabulado.py`,
SQL, columnas, vistas, tolerancias y contenido de sus Excel.

## Fase 4 — división interna del Revisor

Hacerla al final, cuando las utilidades compartidas ya hayan salido. Mantener
`Revisor_Relq/Revisor_Reliquidacion.py` como único punto de entrada y como el
archivo literal que localizan los hermanos. Extraer gradualmente:

1. estado y caché;
2. búsqueda/árbol de archivos;
3. adaptadores de lectura Excel y MDB;
4. motores V4…V17 y sus dependencias;
5. generación del traspaso y lanzamiento de procesos.

No cambiar IDs de verificación, nombres de JSON, argumentos, ubicaciones ni el
orden observable de la ventana. Evitar un `utils.py` general y herencia
profunda: ambos obligarían a leer más contexto para entender una operación.
