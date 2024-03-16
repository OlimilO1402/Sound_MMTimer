.486                 ;Create 32 bit code
.model flat, stdcall ;32 bit memory model
option casemap :all  ;Non Case sensitive

.code

start:

Container proc Param1 : DWORD, Param2 : DWORD, Param3 : DWORD, Param4 : DWORD

push eax             ;Pass the Pointer as last Param to the Function
push Param4          ;Pass all 4 specificated Params
push Param3          ;  "   "  "      "         "
push Param2          ;  "   "  "      "         "
push Param1          ;  "   "  "      "         "
push 55555555h       ;Pass the Pointer of the Class (ObjPtr) witch contains the Function being called (Patched at runtime)

mov eax,66666666h    ;Put the Pointer to the Function here
Call eax             ;Call the Function

ret                  ;Exit Function
Container endp

end start