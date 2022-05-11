
#pragma comment(lib, "libyara.lib")
#define HAVE_LIBCRYPTO

#include "yara.h"
#include <stdio.h>
#include <stdlib.h>
#include <io.h>
#include "util.h"
#include <conio.h>

int ruleNo = 0;

bool FileExists(LPCTSTR szPath)
{
	DWORD dwAttrib = GetFileAttributes(szPath);
	bool rv = (dwAttrib != INVALID_FILE_ATTRIBUTES && !(dwAttrib & FILE_ATTRIBUTE_DIRECTORY)) ? true : false;
	return rv;
}

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

void myExit() {
	printf("Press any key to exit...");
	_getch();
	exit(EXIT_FAILURE);
}

void assert_true_rule_file(rule, filename) {                          
                                                                 
    char* buf;                                                          
	size_t sz = read_file(filename, &buf);

	printf("processing rule: %d datFileSz: 0x%x ", ruleNo++, sz);

    if (sz == 0) {                       
      fprintf(stderr, "%s:%d: cannot read file '%s'\n", __FILE__, __LINE__, filename);                            
	  myExit();
    }             

    if (!matches_blob(rule, (uint8_t*) (buf), sz)) {                    
      fprintf(stderr, "%s:%d: rule does not match contents of '%s' (but should)\n\n", __FILE__, __LINE__, filename);    
	  fprintf(stderr, "%s\n", rule);
	  myExit();
    }                                                                   
    free(buf);                                                          
	
	printf("ok!\n");
 }

int main(int argc, char** argv)
{

  chdir("./tests/");
  printf("Current Working Directory: ");
  system("cd");

  yr_initialize();

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.imports(\"KERNEL32.dll\", \"DeleteCriticalSection\") \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.imports(\"KERNEL32.dll\", \"DeleteCriticalSection\") \
      }",
      "tiny-idata-51ff");

  assert_false_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.imports(\"KERNEL32.dll\", \"DeleteCriticalSection\") \
      }",
      "tiny-idata-5200");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.imports(/.*/, /.*CriticalSection/) \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.imports(/kernel32\\.dll/i, /.*/) \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.imports(/.*/, /.*/) \
      }",
      "tiny-idata-5200");

  assert_false_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.imports(/.*/, /.*CriticalSection/) \
      }",
      "tiny-idata-5200");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.number_of_imports == 2 \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.number_of_sections == 7 \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.entry_point == 0x14E0 \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.linker_version.major == 2 and \
          pe.linker_version.minor == 26 \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.sections[0].name == \".text\" and \
          pe.sections[1].name == \".data\" and \
          pe.sections[2].name == \".rdata\" and \
          pe.sections[3].name == \".bss\" and \
          pe.sections[4].name == \".idata\" and \
          pe.sections[5].name == \".CRT\" and \
          pe.sections[6].name == \".tls\" \
      }",
      "tiny");

  #if defined(HAVE_LIBCRYPTO) || \
      defined(HAVE_WINCRYPT_H) || \
      defined(HAVE_COMMONCRYPTO_COMMONCRYPTO_H)

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.imphash() == \"1720bf764274b7a4052bbef0a71adc0d\" \
      }",
      "tiny");

  #endif

  #if defined(HAVE_LIBCRYPTO)

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.number_of_signatures == 1 and \
          pe.signatures[0].thumbprint == \"c1bf1b8f751bf97626ed77f755f0a393106f2454\" and \
          pe.signatures[0].subject == \"/C=US/ST=California/L=Menlo Park/O=Quicken, Inc./OU=Operations/CN=Quicken, Inc.\" \
      }",
      "079a472d22290a94ebb212aa8015cdc8dd28a968c6b4d3b88acdd58ce2d3b885");

  #endif

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.section_index(\".text\") == 0 \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.section_index(pe.entry_point) == 0 \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.is_32bit() and not pe.is_64bit() \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.checksum == 0xA8DC \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.checksum == pe.calculate_checksum() \
      }",
      "tiny");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.overlay.offset == 0x8000 and pe.overlay.size == 7 \
      }",
      "tiny-overlay");

  assert_true_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
         pe.overlay.size == 0 \
      }",
      "tiny");

  assert_false_rule_file(
      "import \"pe\" \
      rule test { \
        condition: \
          pe.checksum == pe.calculate_checksum() \
      }",
      "tiny-idata-51ff");

  yr_finalize();
  return 0;
}
