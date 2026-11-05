Attribute VB_Name = "modConfig"
' =====================================================================
' modConfig - Shared constants of the Byte Bakers simulator
' Owner: Adriana  |  Task: #2 (A2)
'
' This module is the frozen contract shared by the whole team:
' memory map, sizes, register indexes, sheet names, colors and
' shape name prefixes. Only constants live here (no state, no logic).
' =====================================================================
Option Explicit

' ---------------------------------------------------------------------
' Sizes
' ---------------------------------------------------------------------
Public Const PAGE_SIZE As Integer = 32       ' bytes per page and per frame
Public Const NUM_VPAGES As Integer = 16      ' virtual pages 0-15
Public Const NUM_FRAMES As Integer = 8       ' physical frames 0-7
Public Const VMEM_SIZE As Integer = 512      ' 16 x 32 B virtual memory (swap)
Public Const RAM_SIZE As Integer = 256       ' 8 x 32 B physical memory
Public Const LINE_SIZE As Integer = 8        ' bytes per cache line
Public Const L1_LINES As Integer = 1         ' L1I and L1D: 1 line of 8 B each
Public Const L2_LINES As Integer = 4         ' L2: 4 lines of 8 B = 32 B
Public Const INSTR_SIZE As Integer = 2       ' every instruction is 2 bytes

' ---------------------------------------------------------------------
' Address formats
'   Virtual  (9 bits): [ VPN 4 bits | offset 5 bits ]  VPN = vaddr \ 32
'   Physical (8 bits): [ PFN 3 bits | offset 5 bits ]  paddr = PFN * 32 + offset
' ---------------------------------------------------------------------
Public Const VADDR_BITS As Integer = 9        ' internal address bus, PC and MAR
Public Const PADDR_BITS As Integer = 8        ' external address bus (after the MMU)
Public Const DATA_BITS As Integer = 8         ' data bus and general registers
Public Const OFFSET_BITS As Integer = 5       ' offset inside a 32 B page
Public Const VADDR_MAX As Integer = 511       ' &H1FF
Public Const PADDR_MAX As Integer = 255       ' &HFF
Public Const BYTE_MAX As Integer = 255

' ---------------------------------------------------------------------
' I/O page: page 15 always lives in frame 7.
' It is never evicted and never cached (otherwise the LCD would not refresh).
' ---------------------------------------------------------------------
Public Const IO_VPAGE As Integer = 15
Public Const IO_FRAME As Integer = 7

' Physical I/O addresses (inside frame 7)
Public Const KBD_HEAD As Integer = &HE0       ' keyboard queue head index
Public Const KBD_TAIL As Integer = &HE1       ' keyboard queue tail index
Public Const LCD_CURSOR As Integer = &HE2     ' LCD cursor position 0-7
Public Const KBD_COUNT As Integer = &HE3      ' bytes waiting in the queue 0-8
Public Const KBD_BASE As Integer = &HF0       ' keyboard queue &HF0-&HF7
Public Const LCD_BASE As Integer = &HF8       ' LCD video buffer &HF8-&HFF
Public Const IO_BUF_SIZE As Integer = 8       ' 8 B queue and 8 B video buffer

' Same I/O cells seen from the virtual address space (page 15)
Public Const V_KBD_HEAD As Integer = &H1E0
Public Const V_KBD_TAIL As Integer = &H1E1
Public Const V_LCD_CURSOR As Integer = &H1E2
Public Const V_KBD_COUNT As Integer = &H1E3
Public Const V_KBD_BASE As Integer = &H1F0
Public Const V_LCD_BASE As Integer = &H1F8

' ---------------------------------------------------------------------
' General purpose registers (index inside CPU.R)
' ---------------------------------------------------------------------
Public Const R_AX As Integer = 0              ' accumulator
Public Const R_BX As Integer = 1              ' base
Public Const R_CX As Integer = 2              ' loop counter
Public Const R_DX As Integer = 3              ' data and I/O
Public Const NUM_REGS As Integer = 4

' ---------------------------------------------------------------------
' Page replacement policy
' ---------------------------------------------------------------------
Public Const POLICY_FIFO As Integer = 0
Public Const POLICY_LRU As Integer = 1

' ---------------------------------------------------------------------
' Sheet names
' ---------------------------------------------------------------------
Public Const SH_SIM As String = "Simulator"
Public Const SH_MEM As String = "Memory"
Public Const SH_PROG As String = "Programs"
Public Const SH_LOG As String = "Log"
Public Const SH_ISA As String = "ISA"         ' reference only: cells, no animated shapes

' ---------------------------------------------------------------------
' Colors: hexadecimal BGR Const values (&HBBGGRR&).
' They are constants, so they are never lost when VBA resets.
' ---------------------------------------------------------------------
Public Const CLR_ACTIVE As Long = &H305AD8&    ' RGB(216, 90, 48)   coral: active driver
Public Const CLR_LOAD As Long = &H279FEF&      ' RGB(239, 159, 39)  amber: receiver (Ld_*)
Public Const CLR_HIGHZ As Long = &HA9B2B4&     ' RGB(180, 178, 169) gray: High-Z
Public Const CLR_HIT As Long = &H759E1D&       ' RGB(29, 158, 117)  green: cache hit
Public Const CLR_MISS As Long = &H4A4BE2&       ' RGB(226, 75, 74)   red: cache miss
Public Const CLR_FAULT As Long = &H2D2DA3&     ' RGB(163, 45, 45)   dark red: page fault
Public Const CLR_BG As Long = &H37291F&        ' RGB(31, 41, 55)    dark blue: background
Public Const CLR_TEXT As Long = &HE8EFF1&      ' RGB(241, 239, 232) off-white: text
Public Const CLR_LCD As Long = &H0FBC9B&        ' RGB(155, 188, 15)  LCD green

' ---------------------------------------------------------------------
' Shape name prefixes (used by the animation engine and every panel)
' ---------------------------------------------------------------------
Public Const PFX_REG As String = "reg_"       ' reg_AX       register box
Public Const PFX_BUF As String = "buf_"       ' buf_AX       tri-state buffer
Public Const PFX_WIRE As String = "wire_"     ' wire_DATA    bus line
Public Const PFX_SIG As String = "sig_"       ' sig_Mem_R    control word light
Public Const PFX_L1I As String = "l1i_"       ' l1i_line0    L1I cache line
Public Const PFX_L1D As String = "l1d_"       ' l1d_line0    L1D cache line
Public Const PFX_L2 As String = "l2_"         ' l2_line2     L2 cache line
Public Const PFX_FRM As String = "frm_"       ' frm_3        physical frame
Public Const PFX_PTE As String = "pte_"       ' pte_12       page table entry
Public Const PFX_KBD As String = "kbd_"       ' kbd_q0       keyboard queue cell
Public Const PFX_LCD As String = "lcd_"       ' lcd_0        LCD cell
Public Const PFX_BTN As String = "btn_"       ' btn_Step     button
