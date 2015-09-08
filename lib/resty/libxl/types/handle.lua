local ffi      = require "ffi"
local ffi_cdef = ffi.cdef

ffi_cdef[[
typedef struct tagBookHandle * BookHandle;
typedef struct tagSheetHandle * SheetHandle;
typedef struct tagFormatHandle * FormatHandle;
typedef struct tagFontHandle * FontHandle;
]]