local lib     = require "resty.libxl.library"
local font    = require "resty.libxl.font"
local ffi     = require "ffi"
local C       = ffi.C
local ffi_new = ffi.new
local ffi_str = ffi.string
local fonts   = {}

local s = ffi_new("int[1]", 0)

function fonts.new(opts)
    return setmetatable(opts, fonts)
end

function fonts:add(name, init)
    lib.xlBookAddFontA(self.book.context, name, init and init.context or init)
    return self[self.size]
end

function fonts:__len()
    return lib.xlBookFontSizeA(self.book.context)
end

function fonts:__index(n)
    if n == "size" or n == "count" then
        return lib.xlBookFontSizeA(self.book.context)
    elseif n == "default" then
        return { name = ffi_str(lib.xlBookDefaultFontA(self.book.context, s)), size = s[0] }
    elseif type(n) == "number" then
        return font.new{ context = lib.xlBookFontA(self.book.context, n - 1) }
    else
        return rawget(fonts, n)
    end
end

return fonts