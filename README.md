# Byte Bakers — Simulador Avanzado de Computadora

<p align="center">
  <img src="docs/logo/logo-horizontal.png" alt="Byte Bakers" width="520"/>
</p>

> 🇬🇧 [English version](README_EN.md)

Simulador de una computadora de 8 bits construido en **Excel + VBA**: CPU con 4 registros de propósito general, buses del sistema con **buffers tri-estado** y alta impedancia (High-Z), palabra de control animada, jerarquía de caché **L1I / L1D / L2**, **memoria virtual paginada** de 512 bytes sobre 256 bytes de RAM, y periféricos mapeados en memoria (**teclado** y **pantalla LCD 8x1**).

**Materia:** Arquitectura de Computadoras (SIS-131) — Universidad Católica Boliviana "San Pablo", Santa Cruz
**Evaluación:** Segundo parcial — Proyecto de ingeniería y sustentación oral

---

## Equipo

| Integrante | Rol | Módulos VBA |
|---|---|---|
| **Adriana** | CPU, buses del sistema e integración | `modConfig`, `modState`, `modISA`, `modCPU`, `modALU`, `modBus`, `modControlWord`, `modControlUnit`, `modPanelCPU` |
| **Karen** | Jerarquía de memoria y herramientas del repositorio | `modDevTools`, `modVMem`, `modMMU`, `modMemory`, `modCache`, `modPanelMemory` |
| **Misael** | Entrada/Salida, software e interfaz | `modUI`, `modDashboard`, `modKeyboard`, `modLCD`, `modIO`, `modAssembler`, `modLoader` |

Tablero Kanban: [GitHub Projects](https://github.com/users/Adriana-HS/projects/4)

---

## Índice

| # | Sección | Responsable | Tarea |
|---|---|---|---|
| 1 | [Descarga y ejecución](#1-descarga-y-ejecución) | Adriana | [#11](https://github.com/Adriana-HS/byte-bakers-simulator/issues/11) |
| 2 | [Arquitectura del CPU](#2-arquitectura-del-cpu) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 3 | [Conjunto de instrucciones (ISA)](#3-conjunto-de-instrucciones-isa) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 4 | [Formato de instrucción](#4-formato-de-instrucción) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 5 | [Buses del sistema y buffers tri-estado](#5-buses-del-sistema-y-buffers-tri-estado) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 6 | [Palabra de control](#6-palabra-de-control) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 7 | [Ciclo de instrucción](#7-ciclo-de-instrucción) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 8 | [Mapa de memoria](#8-mapa-de-memoria) | Karen | [#23](https://github.com/Adriana-HS/byte-bakers-simulator/issues/23) |
| 9 | [Memoria virtual y paginación](#9-memoria-virtual-y-paginación) | Karen | [#23](https://github.com/Adriana-HS/byte-bakers-simulator/issues/23) |
| 10 | [Jerarquía de caché L1I / L1D / L2](#10-jerarquía-de-caché-l1i--l1d--l2) | Karen | [#23](https://github.com/Adriana-HS/byte-bakers-simulator/issues/23) |
| 11 | [Entrada/Salida: teclado y LCD 8x1](#11-entradasalida-teclado-y-lcd-8x1) | Misael | [#33](https://github.com/Adriana-HS/byte-bakers-simulator/issues/33) |
| 12 | [Ensamblador, editor y cargador](#12-ensamblador-editor-y-cargador) | Misael | [#33](https://github.com/Adriana-HS/byte-bakers-simulator/issues/33) |
| 13 | [Manual de usuario](#13-manual-de-usuario) | Misael | [#33](https://github.com/Adriana-HS/byte-bakers-simulator/issues/33) |
| 14 | [Metodología de trabajo](#14-metodología-de-trabajo) | Misael | [#33](https://github.com/Adriana-HS/byte-bakers-simulator/issues/33) |
| 15 | [Estructura del repositorio](#15-estructura-del-repositorio) | Adriana | [#1](https://github.com/Adriana-HS/byte-bakers-simulator/issues/1) |
| 16 | [Cómo contribuir](#16-cómo-contribuir) | Karen | [#12](https://github.com/Adriana-HS/byte-bakers-simulator/issues/12) |

---

## 1. Descarga y ejecución

_Pendiente — enlace público al libro `ByteBakers.xlsm` (release v1.0.0)._

## 2. Arquitectura del CPU

CPU de **8 bits** con cuatro registros de propósito general (**AX, BX, CX, DX**), registros de control (**PC, IR, MAR, MDR, PADDR**) y temporarios de ALU (**TA, TB**). Flags: **ZF** (cero), **CF** (acarreo/préstamo), **SF** (signo, bit 7).

| Bloque | Rol |
|---|---|
| Banco de registros | Lectura/escritura con wrap de 8 bits (`0…255`) |
| ALU | ADD/SUB/INC/DEC/AND/OR/XOR/CMP sobre TA/TB; actualiza flags según la ISA |
| Unidad de control | Micro-pasos: fetch F1–F6, decode, execute; un clic = un micro-paso |
| Buses | Dirección interna 9 bits, dirección externa 8 bits, datos 8 bits, High-Z |
| Palabra de control | 27 señales `En_*` / `Ld_*` / Mem / IO + pulso CLK |

Detalle de la hoja `ISA` y del panel Simulator: módulos `modISA`, `modPanelCPU`, `modControlUnit`.

## 3. Conjunto de instrucciones (ISA)

Fuente de verdad: [`src/modISA.bas`](src/modISA.bas) y [`docs/ISA.md`](docs/ISA.md). Resumen:

| Opcode | Mnemonic | Formato | Flags | Descripción |
|---|---|---|---|---|
| `0x00` | NOP | — | - | No hace nada |
| `0x01` | MOV | RI | - | r ← imm |
| `0x02` | MOV | RR | - | r ← r2 |
| `0x03` | LOAD | RM | - | r ← M[addr] |
| `0x04` | STORE | MR | - | M[addr] ← r |
| `0x05`/`0x06` | ADD | RR / RI | Z C S | r ← r + op |
| `0x07`/`0x08` | SUB | RR / RI | Z C S | r ← r − op |
| `0x09`/`0x0A` | INC / DEC | R | Z S | r ← r ± 1 |
| `0x0B`/`0x0C` | CMP | RR / RI | Z C S | resta solo flags |
| `0x0D`–`0x0F` | AND / OR / XOR | RR | Z S | lógica bit a bit |
| `0x10`–`0x14` | JMP / JZ / JNZ / JC / JS | A | - | saltos (condicionales usan flags) |
| `0x15` / `0x16` | IN / OUT | RK / LR | - | teclado / LCD |
| `0x1F` | HLT | — | - | detiene Run |

## 4. Formato de instrucción

Instrucción fija de **2 bytes**:

```
Byte 0: [ opcode: 5 | reg: 2 | A8: 1 ]
Byte 1: [ operand: 8 ]
```

- **reg** = destino (origen en STORE/OUT).
- **A8 + operand** = dirección virtual de 9 bits (`0x000`–`0x1FF`) o inmediato 0–255.
- En formatos RR el operando bajo lleva el registro fuente en los 2 bits bajos.

## 5. Buses del sistema y buffers tri-estado

| Bus | Ancho | Quién conduce | Uso |
|---|---|---|---|
| Dirección **interna** | 9 bits | PC o campo addr de IR (`En_PC` / `En_IR_A`) | PC/IR → MAR (virtual) |
| Dirección **externa** | 8 bits | MMU (PADDR) | acceso a caché / RAM / I/O |
| Datos | 8 bits | un solo `En_*` a la vez | regs ↔ MDR ↔ memoria / KBD / LCD |
| Control | 27 bits + CLK | Unidad de control | señales del micro-paso |

Sin conductor activo el bus vale **High-Z** (`-1`) y el buffer se pinta gris. Un segundo conductor en el mismo micro-paso lanza *Bus collision*.

```mermaid
flowchart LR
  subgraph CPU
    PC --> IntAddr[Bus dir. interna 9b]
    IR --> IntAddr
    IntAddr --> MAR
    REGS[AX..DX] --> DataBus[Bus datos 8b]
    MDR --> DataBus
    ALU --> DataBus
    DataBus --> REGS
    DataBus --> MDR
  end
  MAR --> MMU --> ExtAddr[Bus dir. externa 8b]
  ExtAddr --> MEM[Caché / RAM / IO]
  MEM --> DataBus
  KBD --> DataBus
  DataBus --> LCD
```

## 6. Palabra de control

27 señales (bit 0 = `En_PC` … bit 26 = `En_IR_D`). Tabla completa en la hoja `ISA` (generada por `PaintControlWordOnISA`) y en `modControlWord`.

**Reglas de validación** (cada micro-paso, `ValidateControlWord`):

1. Como máximo **un** conductor de datos: entre `En_MDR`, `En_IR_D`, `En_AX`…`En_DX`, `En_ALU`, `En_KBD`.
2. Como máximo **un** conductor de dirección: entre `En_PC`, `En_IR_A`.
3. Como máximo **un** acceso externo: `Mem_R + Mem_W + IO_R + IO_W ≤ 1`.

Varios `Ld_*` pueden estar activos a la vez (varios receptores del mismo valor).

## 7. Ciclo de instrucción

```mermaid
flowchart LR
  F1[F1 PC to MAR] --> F2[F2 M to MDR]
  F2 --> F3[F3 MDR to IR hi]
  F3 --> F4[F4 PC to MAR]
  F4 --> F5[F5 M to MDR]
  F5 --> F6[F6 MDR to IR lo]
  F6 --> D[Decode]
  D --> E[Execute]
  E --> F1
```

- **STEP** = un micro-paso (`StepMicro`).
- **RUN** = bucle a la velocidad del slider hasta **HLT** o Pause.
- Cada micro-paso escribe una fila en la hoja **Log** (reloj, palabra de control, transferencia).

## 8. Mapa de memoria

El CPU trabaja con **direcciones virtuales** de 9 bits; la **MMU** las traduce a **direcciones físicas** de 8 bits antes de tocar la caché o la RAM.

| Espacio | Ancho | Rango | Tamaño | Formato de la dirección |
|---|---|---|---|---|
| Virtual (PC, MAR) | 9 bits | `0x000`–`0x1FF` | 512 B = 16 páginas × 32 B | `[ VPN: 4 bits \| offset: 5 bits ]` |
| Física (PADDR, RAM) | 8 bits | `0x00`–`0xFF` | 256 B = 8 marcos × 32 B | `[ PFN: 3 bits \| offset: 5 bits ]` |

**Fórmulas de traducción:**

```
VPN    = vaddr \ 32        (división entera)
offset = vaddr Mod 32
paddr  = PFN * 32 + offset
```

Ejemplo: `vaddr = 0x0A3` (163) → `VPN = 5`, `offset = 3`. Si la página 5 está en el marco 2, entonces `paddr = 2 * 32 + 3 = 67 = 0x43`.

**Reparto de páginas y marcos:**

| Páginas virtuales | Marcos físicos | Comentario |
|---|---|---|
| 0–14 | 0–6 | Se paginan: 15 páginas compiten por 7 marcos, por eso hay Page Faults con reemplazo |
| 15 | 7 (fijo) | Página de **E/S**: siempre presente (`Pinned`), nunca se desaloja y **nunca se cachea** |

**Celdas de la página de E/S** (marco 7):

| Celda | Dirección física | Dirección virtual |
|---|---|---|
| `KBD_HEAD` | `0xE0` | `0x1E0` |
| `KBD_TAIL` | `0xE1` | `0x1E1` |
| `LCD_CURSOR` | `0xE2` | `0x1E2` |
| `KBD_COUNT` | `0xE3` | `0x1E3` |
| Cola de teclado (8 B) | `0xF0`–`0xF7` | `0x1F0`–`0x1F7` |
| Video LCD (8 B) | `0xF8`–`0xFF` | `0x1F8`–`0x1FF` |

## 9. Memoria virtual y paginación

La memoria virtual (`VMem`, 512 B) hace de **swap**; la RAM (256 B) guarda solo las páginas que están cargadas en algún marco. La **tabla de páginas** tiene una entrada por página (`PageTable(0..15)`) y `FrameOwner(0..7)` dice qué página ocupa cada marco (`-1` = libre).

| Campo de la entrada | Para qué sirve |
|---|---|
| `Present` | la página está en un marco |
| `PFN` | número de marco donde está |
| `Dirty` | se escribió en RAM desde que se cargó |
| `Ref` | se usó recientemente |
| `LoadedAt` | reloj al cargarla (lo usa **FIFO**) |
| `LastUse` | reloj del último acceso (lo usa **LRU**) |
| `Pinned` | página fija (solo la 15) |

**Flujo de un acceso y del Page Fault:**

```mermaid
flowchart TD
  A[CPU emite vaddr de 9 bits] --> B[VPN = vaddr div 32, offset = vaddr mod 32]
  B --> C{Present en la tabla de paginas?}
  C -- Si --> D[paddr = PFN x 32 + offset]
  C -- No --> E[PAGE FAULT]
  E --> F{Hay marco libre?}
  F -- Si --> J[Cargar la pagina desde VMem al marco]
  F -- No --> G[Elegir victima segun politica, nunca una pagina Pinned]
  G --> H{Victima Dirty?}
  H -- Si --> I[Guardar el marco en el swap VMem]
  H -- No --> K[Descartar sin guardar]
  I --> L[Invalidar las lineas del marco en L1I, L1D y L2]
  K --> L
  L --> J
  J --> M[Actualizar la tabla: Present, PFN, LoadedAt]
  M --> D
```

**Políticas de reemplazo** (se cambian con el selector de la hoja `Memory`):

| Política | Víctima | Característica |
|---|---|---|
| **FIFO** | la página con `LoadedAt` más antiguo | simple, pero puede desalojar una página muy usada (por ejemplo la del código) y sufre la anomalía de Belady |
| **LRU** | la página con `LastUse` más antiguo | conserva lo que se usa seguido; no siempre gana (por ejemplo, un bucle secuencial sobre más páginas que marcos) |

El reloj (`Clock`) suma 1 por cada micro-paso, de modo que sirve tanto para FIFO como para LRU. La página 15 nunca es víctima.

## 10. Jerarquía de caché L1I / L1D / L2

Las cachés trabajan con **direcciones físicas** (después de la MMU). Cada línea guarda 8 bytes y su `Tag = paddr \ 8`.

| Caché | Tamaño | Organización | Quién la usa |
|---|---|---|---|
| **L1I** | 8 B (1 línea) | mapeo directo | solo el fetch de instrucciones |
| **L1D** | 8 B (1 línea) | mapeo directo | solo `LOAD` / `STORE` |
| **L2** | 32 B (4 líneas) | totalmente asociativa, reemplazo **LRU** | unificada (instrucciones y datos) |

```mermaid
flowchart LR
  FETCH[Fetch de instruccion] --> L1I[L1I 8 B]
  DATA[LOAD / STORE] --> L1D[L1D 8 B]
  L1I -- miss --> L2[L2 32 B, 4 lineas, LRU]
  L1D -- miss --> L2
  L2 -- miss --> RAM[RAM 256 B]
  RAM -- llena la linea --> L2
  L2 -- llena la linea --> L1I
  L2 -- llena la linea --> L1D
```

**Lecturas:** se busca en L1; si falla, en L2; si también falla, se trae la línea de 8 bytes desde la RAM, se guarda en L2 y luego en L1. Como cada instrucción mide 2 bytes y una línea guarda 8, **4 instrucciones seguidas dan 1 miss + 7 hits** en L1I (el fetch es byte a byte).

**Política de escritura: write-through + no-write-allocate**

- **Write-through:** cada escritura actualiza también la RAM, así que la RAM siempre está al día (costo: más tráfico hacia la RAM).
- **No-write-allocate:** un `STORE` que falla **no** trae la línea a la caché; solo escribe en RAM.
- Si la línea ya estaba en L1D o en L2, esa copia también se actualiza.
- Las escrituras también cuentan como hit o miss en las estadísticas.

**Excepciones y coherencia:**

- La página 15 (E/S) **nunca se cachea**: `MemRead` y `MemWrite` la saltan.
- Cuando un Page Fault desaloja un marco, `InvalidateFrame(pfn)` invalida las líneas de ese marco en L1I, L1D y L2.
- El editor de RAM de la hoja `Memory` también llama a `InvalidateFrame`, para que la caché no devuelva un dato viejo.

**Estadísticas** (`Stats`): aciertos y fallos de L1I, L1D y L2, y total de Page Faults; se muestran en la hoja `Memory` y en el panel MEMORY del Simulator.

## 11. Entrada/Salida: teclado y LCD 8x1

_Pendiente._

## 12. Ensamblador, editor y cargador

_Pendiente._

## 13. Manual de usuario

El libro es `excel/ByteBakers.xlsm`. Ábrelo siempre desde esa carpeta del repo. El enlace público queda en la release v1.0.0.

### Abrir

1. Abre `excel/ByteBakers.xlsm`.
2. Si sale la barra amarilla, pulsa **Habilitar contenido**.
3. Las hojas están abajo: **Simulator**, **Memory**, **Programs**, **Log**, **ISA**.

### Simulator

![Panel del simulador](docs/img/manual-simulator.png)

- **STEP** avanza un micro-paso. Arriba a la derecha se lee la fase (`FETCH`, `DECODE` o `EXECUTE`), el paso real (`F3  MDR -> IR, PC+1`) y `CLK n`.
- **RUN** sigue hasta `HLT` o hasta **PAUSE**.
- **RESET** pone la máquina en cero y vacía el Log.
- **DEMO** recorre un programa corto.
- Debajo de los botones está la velocidad: lento, normal, rápido o instantáneo.
- El teclado de la izquierda escribe en la cola. **POP** saca una tecla. El LCD muestra lo que el programa manda con `OUT`.
- Los ocho marcos de la derecha son un resumen. La edición se hace en la hoja **Memory**.

### Cargar un programa

![Catálogo y editor](docs/img/manual-programs.png)

1. Ve a **Programs**.
2. **USE 1** carga la multiplicación, **USE 2** el eco del teclado y **USE 3** los page faults. El fuente baja al editor.
3. **VERIFY**. El estado tiene que decir `OK` y cuántos bytes salieron.
4. **LOAD**. El programa queda en la memoria virtual y el PC vuelve a 0.
5. Vuelve a **Simulator** y usa **STEP** o **RUN**.

**SAVE** guarda el texto que editaste. **SAVED** lo vuelve a abrir.

### Teclado

Carga el programa 2, pulsa **RUN** y, mientras espera, toca letras en el teclado de **Simulator**. Cada tecla sale en el LCD. El programa las lee con `IN AX, [KBD]`.

### Log

![Log de micro-pasos](docs/img/manual-log.png)

Cada **STEP** agrega una fila. El título cambia con el número de paso, la fase y el estado (`OK`, `HALT` o `FAULT`). FETCH va en coral, DECODE en ámbar y EXECUTE en verde. A la derecha están los buses, las señales y la caché (verde si acertó, rojo si falló).

### ISA

![Tabla ISA y el botón de colisión](docs/img/manual-isa.png)

La tabla lista cada instrucción y pinta la familia con su color. **FORCE COLLISION** enciende a propósito dos drivers del bus de datos (`En_AX` y `En_BX`) y muestra el error de validación.

### Memory

![Tabla de páginas y caché](docs/img/manual-memory.png)

La tabla de páginas está en `A1`. Un clic en una fila marca el marco y la traducción. Abajo están las líneas de caché y las barras de acierto. FIFO / LRU y HEX / BIN / DEC cambian la vista. El byte que escribas en hexadecimal en la grilla queda en RAM.

## 14. Metodología de trabajo

Una issue, una rama `feature/<persona>/<tema>`, un PR hacia `develop`. El merge es un merge commit, no squash. `main` solo recibe `develop` en la release.

```mermaid
flowchart TD
  feature["feature/persona/tema"] -->|merge commit| develop
  develop -->|release| main
  main --> tag["tag v1.0.0"]
```

El tablero no se salta: **Backlog → To Do → In Progress → In Review → Done**. Done queda después del merge real.

El mensaje de commit es `tipo(modulo): que hace (#issue)`. Ejemplos: `feat(loader): ... (#31)`, `docs(readme): ... (#33)`. En una rama de feature no se commitea el `.xlsm`.

| Área | Persona |
|---|---|
| CPU, ISA, buses, palabra de control, ciclo y release | Adriana |
| Memoria, paginación, caché y programa 3 | Karen |
| Teclado, LCD, ensamblador, programas 1 y 2, manual y marca | Misael |

## 15. Estructura del repositorio

```
byte-bakers-simulator/
├── src/            módulos VBA exportados (.bas)
├── excel/          libro ByteBakers.xlsm
├── programs/       programas del catálogo (.asm)
├── docs/
│   ├── img/        capturas y diagramas
│   ├── logo/       logo en PNG y SVG
│   └── slides/     presentación de la defensa (PDF)
├── .github/        plantillas de issues y pull requests
├── README.md       documentación en español
└── README_EN.md    documentación en inglés
```

## 16. Cómo contribuir

### Configuración inicial (una sola vez)

1. Clona el repositorio y abre SIEMPRE `excel/ByteBakers.xlsm` desde la carpeta `excel/` del repo (nunca una copia en otra carpeta).
2. En Excel: Archivo → Opciones → Centro de confianza → Configuración del Centro de confianza → Configuración de macros → activa **Confiar en el acceso al modelo de objetos de proyectos de VBA**.
3. Configura Git con tu nombre y el correo registrado en tu cuenta de GitHub:
   `git config user.name "Tu Nombre"` y `git config user.email "tu@correo.com"`.
4. Abre el editor de VBA (Alt + F11) y ejecuta **ImportModules** (Alt + F8) para traer a tu libro el código de `src/`.

### Flujo diario

1. `git checkout develop && git pull`
2. Crea tu rama: `git checkout -b feature/<tu-nombre>/<modulo>-<tema>`
3. Abre Excel y ejecuta la macro **ImportModules** (trae el código más reciente de todos desde `src/`).
4. Programa solo tus módulos.
5. Ejecuta la macro **ExportModules** (escribe los módulos en `src/`).
6. `git add src/<tu-modulo>.bas` → `git commit -m "feat(modulo): descripcion (#issue)"` → `git push`
7. Abre un Pull Request hacia `develop` con descripción, cómo probar, captura y `Closes #n`.

### Reglas

- Nunca se hace commit del `.xlsm` (solo quien integra).
- Nadie hace commit directo a `main` ni a `develop`.
- Una tarea = un issue = una rama = un PR. Merge con "Create a merge commit" (no squash).
- Al implementar una función, bórrala del módulo `modStubs_<nombre>` en el mismo PR.
- Al pegar código en el editor VBA: Insertar → Módulo, renombra con F4 y NO pegues la línea `Attribute VB_Name`.
- Haz `git add` solo de tus archivos (nunca `git add .`): al exportar, Excel reescribe también los módulos de los demás con diferencias cosméticas (por ejemplo, quita el `&` final de las constantes hexadecimales). Si aparecen modificados, restáuralos con `git restore src/<modulo>.bas`.