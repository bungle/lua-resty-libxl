local lib          = require "resty.libxl.library"
local setmetatable = setmetatable
local rawget       = rawget
local rawset       = rawset
local ffi          = require "ffi"
local C            = ffi.C
local ffi_new      = ffi.new
local ffi_str      = ffi.string
local type         = type
local hyperlinks   = {}

local rf = ffi_new("int[1]", 0)
local rl = ffi_new("int[1]", 0)
local cf = ffi_new("int[1]", 0)
local cl = ffi_new("int[1]", 0)

function hyperlinks.new(opts)
    return setmetatable(opts, hyperlinks)
end

function hyperlinks:add(hl, rf, rl, cf, cl)
    lib.xlSheetAddHyperlinkA(self.sheet.context, hl, rf, rl, cf, cl)
    return self[self.size]
end

function hyperlinks:del(index)
    if lib.xlSheetDelHyperlinkA(self.sheet.context, index - 1) == 0 then
        return false, self.sheet.book.error
    else
        return true
    end
end

function hyperlinks:__len()
    return lib.xlSheetHyperlinkSizeA(self.sheet.context)
end

function hyperlinks:__index(n)
    if n == "size" or n == "count" or n == "n" then
        return lib.xlSheetHyperlinkSizeA(self.sheet.context)
    elseif type(n) == "number" then
        local name = lib.xlSheetHyperlinkA(self.sheet.context, n - 1, rf, rl, cf, cl)
        if name then
            name = ffi_str(name)
            if name then
                return {
                    index = n,
                    name = name,
                    row = {
                        first = rf,
                        last = rl
                    }, col = {
                        first = cf,
                        last = cl
                    }
                }
            end
        end
        return nil, self.sheet.book.error
    else
        return rawget(hyperlinks, n)
    end
end

return hyperlinks