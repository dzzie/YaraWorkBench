#pragma once

typedef enum {
	cb_output = 0,
	cb_info = 1,
	cb_match = 2,
	cb_update = 3,
	cb_error = 4,
	cb_matchInfo=5,
	cb_moduleInfo=6
} cb_type;

//Public function vb_stdout(ByVal t As cb_type, ByVal lpMsg As Long) as long 
typedef int(__stdcall *vbCallback)(cb_type, char*);
extern vbCallback vbStdOut;

extern void SetVBCallBack(vbCallback cb);

//Public Function GetDebuggerCommand(ByVal buf As Long, ByVal sz As Long) As Long
//typedef int(__stdcall *vbDbgCallback)(char*, int);
//extern vbDbgCallback vbLineInput;
