# Byte Bakers — Advanced Computer Simulator

> 🇪🇸 [Versión en español](README.md)

<!-- Logo: docs/logo/ (task #24) -->

An 8-bit computer simulator built in **Excel + VBA**: a CPU with 4 general-purpose registers, system buses with **tri-state buffers** and high impedance (High-Z), an animated control word, an **L1I / L1D / L2** cache hierarchy, 512 bytes of **paged virtual memory** on top of 256 bytes of RAM, and memory-mapped peripherals (**keyboard** and **8x1 LCD**).

**Course:** Computer Architecture (SIS-131) — Universidad Católica Boliviana "San Pablo", Santa Cruz
**Assessment:** Second midterm — Engineering project and oral defense

---

## Team

| Member | Role | VBA modules |
|---|---|---|
| **Adriana** | CPU, system buses and integration | `modConfig`, `modState`, `modISA`, `modCPU`, `modALU`, `modBus`, `modControlWord`, `modControlUnit`, `modPanelCPU` |
| **Karen** | Memory hierarchy and repository tooling | `modDevTools`, `modVMem`, `modMMU`, `modMemory`, `modCache`, `modPanelMemory` |
| **Misael** | Input/Output, software and user interface | `modUI`, `modDashboard`, `modKeyboard`, `modLCD`, `modIO`, `modAssembler`, `modLoader` |

Kanban board: [GitHub Projects](https://github.com/users/Adriana-HS/projects/4)

---

## Contents

| # | Section | Owner | Task |
|---|---|---|---|
| 1 | [Download and run](#1-download-and-run) | Adriana | [#11](https://github.com/Adriana-HS/byte-bakers-simulator/issues/11) |
| 2 | [CPU architecture](#2-cpu-architecture) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 3 | [Instruction set (ISA)](#3-instruction-set-isa) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 4 | [Instruction format](#4-instruction-format) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 5 | [System buses and tri-state buffers](#5-system-buses-and-tri-state-buffers) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 6 | [Control word](#6-control-word) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 7 | [Instruction cycle](#7-instruction-cycle) | Adriana | [#10](https://github.com/Adriana-HS/byte-bakers-simulator/issues/10) |
| 8 | [Memory map](#8-memory-map) | Karen | [#23](https://github.com/Adriana-HS/byte-bakers-simulator/issues/23) |
| 9 | [Virtual memory and paging](#9-virtual-memory-and-paging) | Karen | [#23](https://github.com/Adriana-HS/byte-bakers-simulator/issues/23) |
| 10 | [L1I / L1D / L2 cache hierarchy](#10-l1i--l1d--l2-cache-hierarchy) | Karen | [#23](https://github.com/Adriana-HS/byte-bakers-simulator/issues/23) |
| 11 | [Input/Output: keyboard and 8x1 LCD](#11-inputoutput-keyboard-and-8x1-lcd) | Misael | [#33](https://github.com/Adriana-HS/byte-bakers-simulator/issues/33) |
| 12 | [Assembler, editor and loader](#12-assembler-editor-and-loader) | Misael | [#33](https://github.com/Adriana-HS/byte-bakers-simulator/issues/33) |
| 13 | [User manual](#13-user-manual) | Misael | [#33](https://github.com/Adriana-HS/byte-bakers-simulator/issues/33) |
| 14 | [Methodology](#14-methodology) | Misael | [#33](https://github.com/Adriana-HS/byte-bakers-simulator/issues/33) |
| 15 | [Repository layout](#15-repository-layout) | Adriana | [#1](https://github.com/Adriana-HS/byte-bakers-simulator/issues/1) |

---

## 1. Download and run

_Pending — public link to the `ByteBakers.xlsm` workbook (release v1.0.0)._

## 2. CPU architecture

_Pending._

## 3. Instruction set (ISA)

_Pending._

## 4. Instruction format

_Pending._

## 5. System buses and tri-state buffers

_Pending._

## 6. Control word

_Pending._

## 7. Instruction cycle

_Pending._

## 8. Memory map

_Pending._

## 9. Virtual memory and paging

_Pending._

## 10. L1I / L1D / L2 cache hierarchy

_Pending._

## 11. Input/Output: keyboard and 8x1 LCD

_Pending._

## 12. Assembler, editor and loader

_Pending._

## 13. User manual

_Pending._

## 14. Methodology

_Pending — Gitflow, Kanban board, commit convention and task split._

## 15. Repository layout

```
byte-bakers-simulator/
├── src/            exported VBA modules (.bas)
├── excel/          ByteBakers.xlsm workbook
├── programs/       catalog programs (.asm)
├── docs/
│   ├── img/        screenshots and diagrams
│   ├── logo/       logo in PNG and SVG
│   └── slides/     defense slide deck (PDF)
├── .github/        issue and pull request templates
├── README.md       Spanish documentation
└── README_EN.md    English documentation
```
