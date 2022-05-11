import "pe"

rule ws32_b
{
        strings:
          $fp = {C6 4? [1-2] 2E C6 4? [1-2] 64 C6 4? [1-2] 6C C6 4? [1-2] 6C C6 4? [1-2] 6B C6 4? [1-2] 65}
          $k32b = {C6 4? [1-2] 6B C6 4? [1-2] 65 C6 4? [1-2] 72 C6 4? [1-2] 6e C6 4? [1-2] 65 C6 4? [1-2] 6c C6 4? [1-2] 33 C6 4? [1-2] 32}
          //$ws32b = {C6 4? [1-2] 77 C6 4? [1-2] 73 C6 4? [1-2] 32 C6 4? [1-2] 5f C6 4? [1-2] 33 C6 4? [1-2] 32} 
        
        condition:
          (uint16(0) == 0x5A4D and uint16(uint32(0x3c)) == 0x4550) and 
          /*pe.number_of_signatures == 0 and  can not use bug in llYara*/
          $k32b and not $fp
}