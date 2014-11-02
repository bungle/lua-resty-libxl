local lib        = require "resty.libxl.library"
local ffi        = require "ffi"
local C          = ffi.C
local ffi_new    = ffi.new
local ffi_str    = ffi.string
local pictures   = {}

local d = ffi_new("const char*[1]")
local l = ffi_new("unsigned[1]", 0)
local s = ffi_new("int[1]", 0)

function pictures.new(opts)
    return setmetatable(opts, pictures)
end

function pictures:add(file, len)
    if len then
        lib.xlBookAddPicture2A(self.book.context, file, len)
    else
        lib.xlBookAddPictureA(self.book.context, file)
    end
    return self[self.size]
end

function pictures:__len()
    return lib.xlBookPictureSizeA(self.book.context)
end

function pictures:__index(n)
    if n == "size" or n == "count" then
        return lib.xlBookPictureSizeA(self.book.context)
    elseif type(n) == "number" then
        local type = lib.xlBookGetPictureA(self.book.context, n - 1, d, l)
        if type == C.PICTURETYPE_ERROR then
            return nil
        else
            return { type = type, data = ffi_str(d[0], l[0]), size = l[0] }
        end
    else
        return rawget(pictures, n)
    end
end

return pictures