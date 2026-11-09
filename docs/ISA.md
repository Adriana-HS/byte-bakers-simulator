# Instruction Set Architecture (ISA)

> Source of truth: [`src/modISA.bas`](../src/modISA.bas). The `ISA` sheet of the workbook is generated from the same table by `BuildISASheet`.

## Instruction format (fixed 2 bytes)

```
Byte 0: [ opcode: 5 bits | reg: 2 bits | A8: 1 bit ]
Byte 1: [ operand: 8 bits ]
```

| Field | Bits | Meaning |
|---|---|---|
| `opcode` | Byte 0, bits 7-3 | Instruction (0-31) |
| `reg` | Byte 0, bits 2-1 | Destination register (source register for `STORE` and `OUT`) |
| `A8` | Byte 0, bit 0 | Bit 8 of a 9-bit address |
| `operand` | Byte 1, bits 7-0 | Immediate 0-255, low byte of an address, or source register in bits 1-0 |

- **Registers:** AX = 0, BX = 1, CX = 2, DX = 3.
- **Addresses** are virtual and 9 bits wide (`0x000`-`0x1FF`): `A8` + `operand`.
- With 2 bytes per instruction, one L1I line (8 B) holds exactly 4 instructions.

## Instructions

Examples use AX as the destination register.

| Opcode | Binary | Mnemonic | Syntax | Format | Flags | Description | Example | Machine code |
|---|---|---|---|---|---|---|---|---|
| `0x00` | `00000` | `NOP` | `NOP` | none | - | Does nothing | `NOP` | `0x0000` |
| `0x01` | `00001` | `MOV` | `MOV r, imm` | RI | - | r ← imm | `MOV AX, 10` | `0x080A` |
| `0x02` | `00010` | `MOV` | `MOV r, r2` | RR | - | r ← r2 | `MOV AX, BX` | `0x1001` |
| `0x03` | `00011` | `LOAD` | `LOAD r, [addr]` | RM | - | r ← M[addr] | `LOAD AX, [0x040]` | `0x1840` |
| `0x04` | `00100` | `STORE` | `STORE [addr], r` | MR | - | M[addr] ← r | `STORE [0x040], AX` | `0x2040` |
| `0x05` | `00101` | `ADD` | `ADD r, r2` | RR | Z C S | r ← r + r2 | `ADD AX, BX` | `0x2801` |
| `0x06` | `00110` | `ADD` | `ADD r, imm` | RI | Z C S | r ← r + imm | `ADD AX, 10` | `0x300A` |
| `0x07` | `00111` | `SUB` | `SUB r, r2` | RR | Z C S | r ← r - r2 | `SUB AX, BX` | `0x3801` |
| `0x08` | `01000` | `SUB` | `SUB r, imm` | RI | Z C S | r ← r - imm | `SUB AX, 10` | `0x400A` |
| `0x09` | `01001` | `INC` | `INC r` | R | Z S | r ← r + 1 | `INC AX` | `0x4800` |
| `0x0A` | `01010` | `DEC` | `DEC r` | R | Z S | r ← r - 1 | `DEC AX` | `0x5000` |
| `0x0B` | `01011` | `CMP` | `CMP r, r2` | RR | Z C S | r - r2 (flags only) | `CMP AX, BX` | `0x5801` |
| `0x0C` | `01100` | `CMP` | `CMP r, imm` | RI | Z C S | r - imm (flags only) | `CMP AX, 10` | `0x600A` |
| `0x0D` | `01101` | `AND` | `AND r, r2` | RR | Z S | r ← r AND r2 | `AND AX, BX` | `0x6801` |
| `0x0E` | `01110` | `OR` | `OR r, r2` | RR | Z S | r ← r OR r2 | `OR AX, BX` | `0x7001` |
| `0x0F` | `01111` | `XOR` | `XOR r, r2` | RR | Z S | r ← r XOR r2 | `XOR AX, BX` | `0x7801` |
| `0x10` | `10000` | `JMP` | `JMP addr` | A | - | PC ← addr | `JMP 0x040` | `0x8040` |
| `0x11` | `10001` | `JZ` | `JZ addr` | A | - | If ZF = 1: PC ← addr | `JZ 0x040` | `0x8840` |
| `0x12` | `10010` | `JNZ` | `JNZ addr` | A | - | If ZF = 0: PC ← addr | `JNZ 0x040` | `0x9040` |
| `0x13` | `10011` | `JC` | `JC addr` | A | - | If CF = 1: PC ← addr | `JC 0x040` | `0x9840` |
| `0x14` | `10100` | `JS` | `JS addr` | A | - | If SF = 1: PC ← addr | `JS 0x040` | `0xA040` |
| `0x15` | `10101` | `IN` | `IN r, [KBD]` | RK | - | r ← keyboard queue head (0 if empty) | `IN AX, [KBD]` | `0xA800` |
| `0x16` | `10110` | `OUT` | `OUT [LCD], r` | LR | - | LCD[cursor] ← r; cursor advances | `OUT [LCD], AX` | `0xB000` |
| `0x1F` | `11111` | `HLT` | `HLT` | none | - | Halts the CPU | `HLT` | `0xF800` |

**Operand formats:** `R` register · `RR` reg, reg · `RI` reg, immediate · `RM` reg, [address] · `MR` [address], reg · `A` address · `RK` reg, [KBD] · `LR` [LCD], reg.
**Flags:** Z = ZF (zero), C = CF (carry/borrow), S = SF (sign, bit 7). `-` = not changed.

## Assembler syntax

- `;` starts a comment.
- Labels end with `:` (`loop:`).
- Numbers: decimal (`25`), hexadecimal (`0x1A`) or character (`'F'`).

## Worked example

`ADD AX, BX` → opcode `0x05`, reg = AX (0), operand = BX (1):

```
Byte 0 = 00101 | 00 | 0 = 0x28
Byte 1 = 00000001       = 0x01      →  IR = 0x2801
```
