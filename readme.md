
<pre>

Author: David Zimmer [dzzie@yahoo.com]
Homepage: http://sandsprite.com/YaraWorkBench/
Installer: http://sandsprite.com/YaraWorkBench/YaraWorkBench_Setup.exe

libyara compiled in:
   release mode 
   as stdcall dll 
   with vs 2017 
   VCRT, openssl & jansson-2.12 all statically linked in (no ext dlls)
   
libyara mods:
  all exports manually set to stdcall (only yr_init/finalize used from vb6)
  pe.c - begin_struct_array("dll_imports"); .name, .funcCount (for dump)
  object.c - yr_object_print_data(vbCallback vbc, cb_printf(), %x in dump if > 9
  object.h - typedef int(__stdcall *vbCallback)(int, char*);
  static IMPORTED_DLL* pe_parse_imports import_errors       

yhelp small helper dll to simplify api usage for vb

helper_test/test.vbp working w/ yhelp.dll 

yara_workbench - gui app with line numbers, syntax highlighting and intellisense for 
                 writing yara signatures with match file offsets, hex editor, and 
                 disassembler  [in development]

The following dlls require at least XP-SP3 (Encode/Decode Pointer)
   capstone.dll, libyara.dll, vbcapstone.dll, yhelp.dll

/direct/fixme.vbp test trying to use only the libyara api direct from vb..
   -> currently crashing <-  

</pre>

![screenshot](https://github.com/dzzie/YaraWorkBench/blob/master/yara_workbench/mainUI.png?raw=true)
