/*
Copyright (c) 2016. The YARA Authors. All Rights Reserved.

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation and/or
other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors
may be used to endorse or promote products derived from this software without
specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR
ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

#include <stdio.h>
#include <io.h>
#include <process.h>

#include <sys/types.h>
#include <sys/stat.h>
#include <fcntl.h>

#include "vb.h"
#include <yara.h>

char compile_error[1024];
int warnings;

extern HANDLE fp;
extern bool dumpModule;

typedef struct
{
	char* expected;
	int found;

} find_string_t;

int logMsg(cb_type c, const char *format, ...)
{
	DWORD dwErr = GetLastError();
	int rv = 0;

	if (format) {
		char buf[1024];
		va_list args;
		va_start(args, format);
		_vsnprintf(buf, 1024, format, args);
		rv = vbStdOut(c,buf);
		va_end(args);
	}

	SetLastError(dwErr);
	return rv;
}

static void callback_function(int error_level, const char* file_name, int line_number, const char* message, void* user_data)
{
	  if (error_level == YARA_ERROR_LEVEL_WARNING) (*((int*) user_data))++;
	  logMsg(cb_error, "line %d: %s", line_number, message);
}


int compile_rule(char* string, YR_RULES** rules)
{
  YR_COMPILER* compiler = NULL;
  int result = ERROR_SUCCESS;

  compile_error[0] = '\0';
  warnings = 0;

  if (yr_compiler_create(&compiler) != ERROR_SUCCESS)
  {
	vbStdOut(cb_error, "yr_compiler_create");
    goto _exit;
  }

  yr_compiler_set_callback(compiler, callback_function, &warnings);

  if (yr_compiler_add_string(compiler, string, NULL) != 0)
  {
    result = compiler->last_error;
    goto _exit;
  }

  result = yr_compiler_get_rules(compiler, rules);

_exit:
  yr_compiler_destroy(compiler);
  return result;
}

int count_matches(int message, void* message_data, void* user_data)
{

	//CALLBACK_MSG_RULE_MATCHING, CALLBACK_MSG_RULE_NOT_MATCHING, CALLBACK_MSG_SCAN_FINISHED, CALLBACK_MSG_IMPORT_MODULE
	if(vbStdOut(cb_update, (char*)message) == -1) return CALLBACK_ABORT;

	if (message == CALLBACK_MSG_RULE_MATCHING) {
		(*(int*)user_data)++;
		YR_RULE* rule = (YR_RULE*)message_data;
		//if (fp != NULL) {logMsg(cb_match, "match %s.%s offset:%x", rule->ns->name, rule->identifier, GetFilePointer(fp));
		
		logMsg(cb_match, "match %s.%s", rule->ns->name, rule->identifier);
		 
		YR_STRING* string;
		yr_rule_strings_foreach(rule, string)
		{
			YR_MATCH* match;
			yr_string_matches_foreach(string, match)
			{
					logMsg(cb_matchInfo, "\t0x%" PRIx64 ":%d:%s",
						match->base + match->offset,
						match->data_length,
						string->identifier);
				 
					/*if (STRING_IS_HEX(string))
						print_hex_string(match->data, match->data_length);
					else
						print_string(match->data, match->data_length);*/
				 
			}
		}

	}

	if (message == CALLBACK_MSG_MODULE_IMPORTED)
	{

		if (dumpModule)
		{
			YR_OBJECT* object = (YR_OBJECT*)message_data;

			//mutex_lock(&output_mutex);

			yr_object_print_data(object, 0, 1); //assumes SetVBCallBack already set in initilization or no output. (less mods on update libYara)
			vbStdOut(cb_moduleInfo,"\r\n");

			//mutex_unlock(&output_mutex);
		}

	}


	return CALLBACK_CONTINUE;
}


/*
DWORD GetFilePointer(HANDLE hFile) {
DWORD rv = SetFilePointer(hFile, 0, NULL, FILE_CURRENT);
return rv; //nice try but doesnt work...
}


char* print_hex_string(const uint8_t* data, int length)
{

for (int i = 0; i < min(32, length); i++)
printf("%s%02X", (i == 0 ? "" : " "), data[i]);

puts(length > 32 ? " ..." : "");
}

static void print_string(const uint8_t* data,int length)
{
for (int i = 0; i < length; i++)
{
if (data[i] >= 32 && data[i] <= 126)
printf("%c", data[i]);
else
printf("\\x%02X", data[i]);
}

printf("\n");
}

*/

/*

int do_nothing(int message, void* message_data, void* user_data)
{
  return CALLBACK_CONTINUE;
}


int matches_blob(char* rule, uint8_t* blob, size_t len)
{
  YR_RULES* rules;
  find_string_t f;

  f.found = 0;
  f.expected = NULL;

  if (blob == NULL) return 0;
  if (compile_rule(rule, &rules) != ERROR_SUCCESS) return 0;
  
  int matches = 0;
  int scan_result = yr_rules_scan_mem(rules, blob, len, 0, count_matches, &matches, 0);


  if (scan_result != ERROR_SUCCESS)
  {
	  vbStdOut(cb_error, "yr_rules_scan_mem: error");
	  return 0;
  }

  yr_rules_destroy(rules);

  return matches;
}

static int capture_matches(int message, void* message_data, void* user_data)
{
	//CALLBACK_MSG_RULE_MATCHING, CALLBACK_MSG_RULE_NOT_MATCHING, CALLBACK_MSG_SCAN_FINISHED, CALLBACK_MSG_IMPORT_MODULE

	if (message == CALLBACK_MSG_RULE_MATCHING)
	{
		find_string_t* f = (find_string_t*)user_data;
		vbStdOut(cb_match, "in capture match rule");

		YR_RULE* rule = (YR_RULE*)message_data;
		YR_STRING* string;

		yr_rule_strings_foreach(rule, string)
		{
			YR_MATCH* match;

			yr_string_matches_foreach(string, match)
			{
				if (strlen(f->expected) == match->data_length &&
					strncmp(f->expected, (char*)(match->data), match->data_length) == 0)
				{
					f->found++;
					snprintf(compile_error, sizeof(compile_error), "match found offset: %x, matchLen: %d ", match->offset, match->match_length);
					vbStdOut(cb_match, compile_error);
				}
			}
		}
	}

	return CALLBACK_CONTINUE;
}

int matches_string(char* rule, char* string)
{
  size_t len = 0;
  if (string != NULL) len = strlen(string);
  return matches_blob(rule, (uint8_t*)string, len);
}

int capture_string(
    char* rule,
    char* string,
    char* expected_string)
{
  YR_RULES* rules;

  if (compile_rule(rule, &rules) != ERROR_SUCCESS) return 0;

  find_string_t f;

  f.found = 0;
  f.expected = expected_string;

  if (yr_rules_scan_mem(rules, (uint8_t*)string, strlen(string), 0, capture_matches, &f, 0) != ERROR_SUCCESS)
	  return 0;

  yr_rules_destroy(rules);

  return f.found;
}
*/

int file_length(FILE *f)
{
	int pos;
	int end;

	pos = ftell(f);
	fseek(f, 0, SEEK_END);
	end = ftell(f);
	fseek(f, pos, SEEK_SET);

	return end;
}

int read_file(char* filename, char** buf) {
	FILE *fp;
	int sz = 0;
	fp = fopen(filename, "rb");
	if (fp == 0) {
		printf("Failed to open file %s\n", filename);
		return 0;
	}
	sz = file_length(fp);
	*buf = (unsigned char*)malloc(sz);
	memset(*buf, 0, sz);
	fread(*buf, 1, sz, fp);
	fclose(fp);
	return sz;
}

bool fileExists(LPCTSTR szPath)
{
	if (szPath == NULL) return false;
	DWORD dwAttrib = GetFileAttributes(szPath);
	bool rv = (dwAttrib != INVALID_FILE_ATTRIBUTES && !(dwAttrib & FILE_ATTRIBUTE_DIRECTORY)) ? true : false;
	return rv;
}

bool fileExistsW(LPCWSTR szPath)
{
	if (szPath == NULL) return false;
	DWORD dwAttrib = GetFileAttributesW(szPath);
	bool rv = (dwAttrib != INVALID_FILE_ATTRIBUTES && !(dwAttrib & FILE_ATTRIBUTE_DIRECTORY)) ? true : false;
	return rv;
}
