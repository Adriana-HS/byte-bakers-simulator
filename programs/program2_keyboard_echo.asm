; Program 2: echo 8 keys from the keyboard onto the LCD.
        MOV  CX, 8          ; eight LCD positions
read:   IN   AX, [KBD]      ; take the next key
        CMP  AX, 0          ; is the queue empty?
        JZ   read           ; yes: wait for a key
        OUT  [LCD], AX      ; show the character
        DEC  CX             ; one position filled
        JNZ  read           ; repeat until CX is 0
        HLT                 ; stop
