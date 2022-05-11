#pragma once

#include <stdarg.h>

#pragma comment(lib, "./../lib/jansson.lib")
#pragma comment(lib, "./../lib/libcrypto.lib")
#pragma comment(lib, "./../lib/libssl.lib")
#pragma comment(lib, "ws2_32.lib")
#pragma comment(lib, "Crypt32.lib")

#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

typedef enum {
	cb_output = 0,
	cb_info = 1,
	cb_match = 2,
	cb_update = 3,
	cb_error = 4,
	cb_matchInfo = 5,
	cb_moduleInfo = 6,
	cb_dbg = 7
} cb_type;

//Public function vb_stdout(ByVal t As cb_type, ByVal lpMsg As Long) as long 
typedef int(__stdcall *vbCallback)(cb_type, char*);

//below all in object.c
extern vbCallback vbStdOut;
extern void vb_dbg(cb_type ct, char* format, ...);
