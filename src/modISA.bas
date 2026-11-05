Attribute VB_Name = "modISA"
' =====================================================================
' modISA - Instruction set of the Byte Bakers CPU
' Owner: Adriana  |  Task: #3 (A3)
'
' Instruction format (fixed 2 bytes):
'   Byte 0: [ opcode: 5 bits | reg: 2 bits | A8: 1 bit ]
'   Byte 1: [ operand: 8 bits ]
'   - Immediate:       operand = value 0-255
'   - Address:         A8 + operand = 9-bit virtual address 0-511
'   - Source register: 2 low bits of the operand (AX=0 BX=1 CX=2 DX=3)
' =====================================================================
Option Explicit

' ---------------------------------------------------------------------
' Opcodes (5 bits: 0-31)
' ---------------------------------------------------------------------
Public Const OP_NOP As Integer = &H0          ' NOP
Public Const OP_MOV_RI As Integer = &H1       ' MOV r, imm
Public Const OP_MOV_RR As Integer = &H2       ' MOV r, r2
Public Const OP_LOAD As Integer = &H3         ' LOAD r, [addr]
Public Const OP_STORE As Integer = &H4        ' STORE [addr], r
Public Const OP_ADD_RR As Integer = &H5       ' ADD r, r2
Public Const OP_ADD_RI As Integer = &H6       ' ADD r, imm
Public Const OP_SUB_RR As Integer = &H7       ' SUB r, r2
Public Const OP_SUB_RI As Integer = &H8       ' SUB r, imm
Public Const OP_INC As Integer = &H9          ' INC r
Public Const OP_DEC As Integer = &HA          ' DEC r
Public Const OP_CMP_RR As Integer = &HB       ' CMP r, r2
Public Const OP_CMP_RI As Integer = &HC       ' CMP r, imm
Public Const OP_AND As Integer = &HD          ' AND r, r2
Public Const OP_OR As Integer = &HE           ' OR r, r2
Public Const OP_XOR As Integer = &HF          ' XOR r, r2
Public Const OP_JMP As Integer = &H10         ' JMP addr
Public Const OP_JZ As Integer = &H11          ' JZ addr
Public Const OP_JNZ As Integer = &H12         ' JNZ addr
Public Const OP_JC As Integer = &H13          ' JC addr
Public Const OP_JS As Integer = &H14          ' JS addr
Public Const OP_IN As Integer = &H15          ' IN r, [KBD]
Public Const OP_OUT As Integer = &H16         ' OUT [LCD], r
Public Const OP_HLT As Integer = &H1F         ' HLT

Public Const OPCODE_MAX As Integer = 31       ' 5-bit opcode field
Public Const NUM_INSTRUCTIONS As Integer = 24 ' opcodes in use
