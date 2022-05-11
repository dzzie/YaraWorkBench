Attribute VB_Name = "mYara"
Option Explicit

'int __stdcall yr_initialize(void);
Declare Function yr_initialize Lib "libyara" () As Long

'int __stdcall yr_finalize(void);
Declare Function yr_finalize Lib "libyara" () As Long


'int __stdcall yr_compiler_create(YR_COMPILER** compiler);
Declare Function yr_compiler_create Lib "libyara" (ByRef compiler As Long) As Long

'void __stdcall yr_compiler_destroy(YR_COMPILER* compiler);
Declare Sub yr_compiler_destroy Lib "libyara" (ByVal compiler As Long)

'YR_API void __stdcall yr_compiler_set_callback(
'    YR_COMPILER* compiler,
'    YR_COMPILER_CALLBACK_FUNC callback,
'    void* user_data);
Declare Sub yr_compiler_set_callback Lib "libyara" (ByVal compiler As Long, ByVal lpCallBack As Long, ByRef userData As Long)


'    YR_API int __stdcall yr_compiler_add_file(
'    YR_COMPILER* compiler,
'    FILE* rules_file,
'    const char* namespace_,
'    const char* file_name);
'
'YR_API int __stdcall yr_compiler_add_string(
'    YR_COMPILER* compiler,
'    const char* rules_string,
'    const char* namespace_);
Declare Function yr_compiler_add_string Lib "libyara" (ByRef compiler As Long, ByVal yara_rule As Long, ByVal namespace As Long)

'YR_API char* __stdcall yr_compiler_get_error_message(
'    YR_COMPILER* compiler,
'    char* buffer,
'    int buffer_size);



'static void __stdcall callback_function(
'    int error_level,
'    const char* file_name,
'    int line_number,
'    const char* message,
'    void* user_data)
'{
'  if (error_level == YARA_ERROR_LEVEL_WARNING)
'    (*((int*) user_data))++;
'
'  snprintf(
'      compile_error,
'      sizeof(compile_error),
'      "line %d: %s",
'      line_number,
'      message);
'}

Public Sub callback(ByVal errLevel As Long, ByVal pFname As Long, ByVal lineNo As Long, ByVal msg As Long, ByVal userData As Long)
    magbox "In compiler callback!"
End Sub


