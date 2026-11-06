# Byte Bakers — Simulador Avanzado de Computadora

> 🇬🇧 [English version](README_EN.md)

<!-- Logo: docs/logo/ (tarea #24) -->

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

_Pendiente._

## 3. Conjunto de instrucciones (ISA)

_Pendiente._

## 4. Formato de instrucción

_Pendiente._

## 5. Buses del sistema y buffers tri-estado

_Pendiente._

## 6. Palabra de control

_Pendiente._

## 7. Ciclo de instrucción

_Pendiente._

## 8. Mapa de memoria

_Pendiente._

## 9. Memoria virtual y paginación

_Pendiente._

## 10. Jerarquía de caché L1I / L1D / L2

_Pendiente._

## 11. Entrada/Salida: teclado y LCD 8x1

_Pendiente._

## 12. Ensamblador, editor y cargador

_Pendiente._

## 13. Manual de usuario

_Pendiente._

## 14. Metodología de trabajo

_Pendiente — Gitflow, tablero Kanban, convención de commits y reparto de tareas._

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