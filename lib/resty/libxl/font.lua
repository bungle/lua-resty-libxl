require "resty.libxl.types.handle"
require "resty.libxl.types.font"
local setmetatable = setmetatable
local rawget       = rawget
local rawset       = rawset
local ffi          = require "ffi"
local ffi_str      = ffi.string
local lib          = require "resty.libxl.library"

local font = {}

function font.new(opts)
    return setmetatable(opts, font)
end

function font:__index(n)
    if n == "size" then
        return lib.xlFontSizeA(self.context)
    elseif n == "italic" then
        return lib.xlFontItalicA(self.context)
    elseif n == "strike" then
        return lib.xlFontStrikeOutA(self.context)
    elseif n == "color" then
        return lib.xlFontColorA(self.context)
    elseif n == "bold" then
        return lib.xlFontBoldA(self.context)
    elseif n == "script" then
        return lib.xlFontScriptA(self.context)
    elseif n == "underline" then
        return lib.xlFontUnderlineA(self.context)
    elseif n == "name" then
        return ffi_str(lib.xlFontNameA(self.context))
    else
        return rawget(font, n)
    end
end

function font:__newindex(n, v)
    if n == "size" then
        lib.xlFontSetSizeA(self.context, v)
    elseif n == "italic" then
        lib.xlFontSetSizeA(self.context, v)
    elseif n == "strike" then
        lib.xlFontSetStrikeOutA(self.context, v)
    elseif n == "color" then
        lib.xlFontSetColorA(self.context, v)
    elseif n == "bold" then
        lib.xlFontSetBoldA(self.context, v)
    elseif n == "script" then
        lib.xlFontSetScriptA(self.context, v)
    elseif n == "underline" then
        lib.xlFontSetUnderlineA(self.context, v)
    elseif n == "name" then
        lib.xlFontSetNameA(self.context, v)
    else
        rawset(self, n, v)
    end
end

return font
