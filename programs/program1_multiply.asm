; Program 1: 7 x 5 by repeated addition.
; AX accumulator, BX multiplicand, CX counter, DX result.
        MOV  AX, 0          ; accumulator starts at 0
        MOV  BX, 7          ; multiplicand
        MOV  CX, 5          ; add BX this many times
mult:   ADD  AX, BX         ; AX = AX + 7
        DEC  CX             ; one pass done
        JNZ  mult           ; repeat while CX is not 0
        MOV  DX, AX         ; DX holds 35
        STORE [0x080], DX   ; keep the result at 0x080
        HLT                 ; stop
