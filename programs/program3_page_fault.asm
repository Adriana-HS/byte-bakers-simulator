; Program 3: scattered accesses to force Page Faults and cache misses
; Code in page 0; data in 8 different pages (1, 3, 5, 7, 9, 11, 13, 14).
; 9 pages in use > 7 available frames -> replacement on every loop.
; Expected (simulated): FIFO approx 28 Page Faults, LRU approx 25.
; To be confirmed in the real simulator once the Clock advances.
        MOV  CX, 3            ; page 0 (code): 3 loops
again:  LOAD AX, [0x020]      ; page 0 (code) + data page 1
        LOAD BX, [0x060]      ; page 0 (code) + data page 3
        ADD  AX, BX           ; page 0 (code), no memory access
        STORE [0x0A0], AX     ; page 0 (code) + data page 5 (dirty -> written back to swap on eviction)
        LOAD DX, [0x0E0]      ; page 0 (code) + data page 7
        STORE [0x120], DX     ; page 0 (code) + data page 9 (dirty)
        LOAD BX, [0x160]      ; page 0 (code) + data page 11
        STORE [0x1A0], BX     ; page 0 (code) + data page 13 (dirty)
        LOAD AX, [0x1C0]      ; page 0 (code) + data page 14
        DEC  CX               ; page 0 (code), no memory access
        JNZ  again            ; page 0 (code): jump back while CX is not 0
        HLT                   ; page 0 (code): stop