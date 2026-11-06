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
