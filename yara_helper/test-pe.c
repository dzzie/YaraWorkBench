
#pragma comment(lib, "./../libyara/libyara.lib")
#define HAVE_LIBCRYPTO

#include <yara.h>
#include <stdio.h>
#include <stdlib.h>
#include <io.h>
#include "util.h"
#include <conio.h>
#include <Windows.h>
#include "vb.h"

#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

vbCallback vbStdOut = 0;
HANDLE fp = 0;
int dumpModule = 0;

extern int logMsg(cb_type c, const char *format, ...);

/*
   yr_scanner_set_timeout(args->scanner, timeout - elapsed_time);
   result = yr_scanner_scan_file(args->scanner, file_path);

*/

typedef enum{
	yo_init = 0,
	yo_term = 1,
	yo_setCallBack = 2,
	yo_getVer = 3,
	yo_InitModuleDump = 4
} yr_opts;

int __stdcall yr_op(yr_opts o, int arg1) {
#pragma EXPORT
	
	if (o == yo_init) return yr_initialize();
	if (o == yo_term) return yr_finalize();
	if (o == yo_getVer && arg1 !=0) return yr_getVersion(arg1);
	
	if (o == yo_InitModuleDump) {
		dumpModule = arg1;
		return 1;
	}

	if (o == yo_setCallBack && arg1 != 0) {
		SetVBCallBack(arg1); //sets the callback in libYara for module dumping for us..
		vbStdOut = (vbCallback)(arg1);
		return 1;
	}

	return 0;

}

//in 3.10 yr_rules_scan_file() was wchar_t* filename,
bool __stdcall testFile(char* rule, char* testString, char* fName) {                          
#pragma EXPORT

	YR_RULES* rules;
	char* buf = 0; 
	char* ruleBuf = 0;
	size_t sz = 0;
	bool rv = false;
	int matches = 0;
	int scan_result = 0;

	if (fileExists(rule)) {
		vbStdOut(cb_info, "loading rule file");
		sz = read_file(rule, &ruleBuf);
		if (sz == 0)  return false;
	}
	else {
		if (rule == NULL) {
			vbStdOut(cb_info, "rule as text can not be null");
			return false;
		}
		vbStdOut(cb_info, "using rule as text");
		ruleBuf = _strdup(rule);
	}

	if (compile_rule(rule, &rules) != ERROR_SUCCESS) return 0;
	
	if (fName!=NULL){
		if (fileExists(fName)) {
			vbStdOut(cb_info, "file scan mode");
			scan_result = yr_rules_scan_file(rules, fName, 0, count_matches, &matches, 0);
			//fp = CreateFile(testString, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
			//scan_result = yr_rules_scan_fd(rules, fp, 0, count_matches, &matches, 0); GetFilePointer doesnt work...
			//CloseHandle(fp);
		}
		else {
			vbStdOut(cb_info, "file scan mode: file not found");
		}
	}
	else {
		scan_result = yr_rules_scan_mem(rules, testString, strlen(testString), 0, count_matches, &matches, 0);
	}

	//file: ERROR_SUCCESS,ERROR_INSUFICENT_MEMORY,ERROR_COULD_NOT_MAP_FILE,ERROR_ZERO_LENGTH_FILE
	//      ERROR_TOO_MANY_SCAN_THREADS,ERROR_SCAN_TIMEOUT,ERROR_CALLBACK_ERROR,ERROR_TOO_MANY_MATCHES

	//mem:  ERROR_SUCCESS, ERROR_INSUFICENT_MEMORY, ERROR_TOO_MANY_SCAN_THREADS, ERROR_SCAN_TIMEOUT, 
	//      ERROR_CALLBACK_ERROR, ERROR_TOO_MANY_MATCHES
	if (scan_result != ERROR_SUCCESS)
	{
		logMsg(cb_error, "testFile error scan_result = %d", scan_result);
		return 0;
	}

	yr_rules_destroy(rules);

    free(buf);                                                          
	free(ruleBuf);
	return matches;
 }
 