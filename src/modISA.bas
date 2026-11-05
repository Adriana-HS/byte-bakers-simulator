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

' =====================================================================
' Encode - packs one instruction into 16 bits.
'   op      opcode 0-31
'   reg     destination register 0-3 (or source register for STORE/OUT)
'   operand immediate 0-255, 9-bit address 0-511 or source register 0-3
' Example: Encode(OP_ADD_RR, R_AX, R_BX) = &H2801  (ADD AX, BX)
' =====================================================================
Public Function Encode(ByVal op As Integer, ByVal reg As Integer, ByVal operand As Integer) As Long
    Dim b0 As Long
    op = op And OPCODE_MAX
    reg = reg And 3
    operand = operand And VADDR_MAX               ' 9 bits: A8 + 8-bit operand
    b0 = CLng(op) * 8 + reg * 2 + (operand \ 256) ' opcode | reg | A8
    Encode = b0 * 256 + (operand And BYTE_MAX)
End Function

' =====================================================================
' Decode - splits the 16-bit IR into its fields (inverse of Encode).
'   operand comes back with A8 included (0-511); for immediates and
'   registers A8 is 0, so the value is the same 0-255.
' =====================================================================
Public Sub Decode(ByVal ir As Long, ByRef op As Integer, ByRef reg As Integer, ByRef operand As Integer)
    Dim b0 As Long
    b0 = (ir \ 256) And BYTE_MAX
    op = b0 \ 8
    reg = (b0 \ 2) And 3
    operand = (b0 And 1) * 256 + (ir And BYTE_MAX)
End Sub
