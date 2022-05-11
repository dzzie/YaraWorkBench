import "pe"

rule isPE
{
  condition:
     // MZ signature at offset 0 and ...
     uint16(0) == 0x5A4D and
     // ... PE signature at offset stored in MZ header at 0x3C
     uint32(uint32(0x3C)) == 0x00004550
}

rule sample
{
        strings:
          	$b = "MSVBVM60"
			$x = {55 8B EC 83 EC 0C} 
	      
        condition:
          	any of them
}