require "resty.libxl.types.handle"
require "resty.libxl.types.book"
require "resty.libxl.types.enum"
local setmetatable = setmetatable
local rawget       = rawget
local rawset       = rawset
local type         = type
local ffi          = require "ffi"
local ffi_gc       = ffi.gc
local ffi_str      = ffi.string
local lib          = require "resty.libxl.library"
local date         = require "resty.libxl.date"
local color        = require "resty.libxl.color"
local sheets       = require "resty.libxl.sheets"
local formats      = require "resty.libxl.formats"
local fonts        = require "resty.libxl.fonts"
local pictures     = require "resty.libxl.pictures"

local book = {}

--[[
APIs not implemented:
        int __cdecl xlBookLoadRawA(BookHandle handle, const char* data, unsigned size);
        int __cdecl xlBookSaveRawA(BookHandle handle, const char** data, unsigned* size);
        int __cdecl xlBookAddCustomNumFormatA(BookHandle handle, const char* customNumFormat);
const char* __cdecl xlBookCustomNumFormatA(BookHandle handle, int fmt);
]]

function book:__index(n)
    if n == "version" then
        return lib.xlBookVersionA(self.context)
    elseif n == "biffversion" then
        return lib.xlBookBiffVersionA(self.context)
    elseif n == "refr1c1" then
        return lib.xlBookRefR1C1A(self.context) == 1
    elseif n == "rgbmode" then
        return lib.xlBookRgbModeA(self.context) == 1
    elseif n == "date1904" then
        return lib.xlBookIsDate1904A(self.context) == 1
    elseif n == "template" then
        return lib.xlBookIsTemplateA(self.context) == 1
    elseif n == "error" then
        return ffi_str(lib.xlBookErrorMessageA(self.context))
    else
        return rawget(book, n)
    end
end

function book:__newindex(n, v)
    if n == "refr1c1" then
        lib.xlBookSetRefR1C1A(self.context, v and 1 or 0)
    elseif n == "rgbmode" then
        lib.xlBookSetRgbModeA(self.context, v and 1 or 0)
    elseif n == "date1904" then
        return lib.xlBookSetDate1904A(self.context, v and 1 or 0)
    elseif n == "template" then
        return lib.xlBookSetTemplateA(self.context, v and 1 or 0)
    elseif n == "locale" then
        return lib.xlBookSetLocaleA(self.context, v)
    else
        rawset(self, n, v)
    end
end

function book.new(opts)
    opts = opts or {}
    local context = opts.format == "xls" and ffi_gc(lib.xlCreateBookCA(), lib.xlBookReleaseA) or  ffi_gc(lib.xlCreateXMLBookCA(), lib.xlBookReleaseA)
    if type(opts.name) == "string" and type(opts.key) == "string" then
        lib.xlBookSetKeyA(context, opts.name, opts.key)
    end
    local self = setmetatable({
        context = context,
           date = date.new{ context = context },
          color = color.new{ context = context }
    }, book)
    self.sheets   = sheets.new{ book = self }
    self.formats  = formats.new{ book = self }
    self.fonts    = fonts.new{ book = self }
    self.pictures = pictures.new{ book = self }
    return self
end

function book:load(filename)
    if lib.xlBookLoadA(self.context, filename) == 1 then
        return true
    else
        return false, self.error
    end
end

function book:save(filename)
    if lib.xlBookSaveA(self.context, filename) == 1 then
        return true
    else
        return false, self.error
    end
end

function book:release()
    return lib.xlBookReleaseA(self.context)
end

return book