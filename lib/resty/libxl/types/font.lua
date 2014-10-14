local ffi        = require "ffi"
local ffi_cdef   = ffi.cdef

ffi_cdef[[
        int __cdecl xlFontSizeA(FontHandle handle);
       void __cdecl xlFontSetSizeA(FontHandle handle, int size);

        int __cdecl xlFontItalicA(FontHandle handle);
       void __cdecl xlFontSetItalicA(FontHandle handle, int italic);

        int __cdecl xlFontStrikeOutA(FontHandle handle);
       void __cdecl xlFontSetStrikeOutA(FontHandle handle, int strikeOut);

        int __cdecl xlFontColorA(FontHandle handle);
       void __cdecl xlFontSetColorA(FontHandle handle, int color);

        int __cdecl xlFontBoldA(FontHandle handle);
       void __cdecl xlFontSetBoldA(FontHandle handle, int bold);

        int __cdecl xlFontScriptA(FontHandle handle);
       void __cdecl xlFontSetScriptA(FontHandle handle, int script);

        int __cdecl xlFontUnderlineA(FontHandle handle);
       void __cdecl xlFontSetUnderlineA(FontHandle handle, int underline);

const char* __cdecl xlFontNameA(FontHandle handle);
       void __cdecl xlFontSetNameA(FontHandle handle, const char* name);
]]