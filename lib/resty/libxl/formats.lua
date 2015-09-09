local lib     = require "resty.libxl.library"
local format  = require "resty.libxl.format"
local rawget  = rawget
local type    = type
local formats = {}

function formats.new(opts)
    return setmetatable(opts, formats)
end

function formats:add(name, init)
    lib.xlBookAddFormatA(self.book.context, name, init and init.context or init)
    return self[self.count]
end

function formats:__len()
    return lib.xlBookFormatSizeA(self.book.context)
end

function formats:__index(n)
    if n == "size" or n == "count" or n == "n" then
        return lib.xlBookFormatSizeA(self.book.context)
    elseif type(n) == "number" then
        return format.new{ index = n, context = lib.xlBookFormatA(self.book.context, n - 1) }
    else
        return rawget(formats, n)
    end
end

return formats