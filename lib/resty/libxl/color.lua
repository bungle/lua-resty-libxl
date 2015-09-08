local lib          = require "resty.libxl.library"
local setmetatable = setmetatable
local ffi          = require "ffi"
local ffi_new      = ffi.new

local r = ffi_new("int[1]", 0)
local g = ffi_new("int[1]", 0)
local b = ffi_new("int[1]", 0)

local color   = {}
color.__index = color

function color.new(opts)
    return setmetatable(opts, color)
end

function color:pack(r, g, b)
    return lib.xlBookColorPackA(self.context, r or 0, g or 0, b or 0)
end

function color:unpack(value)
    lib.xlBookColorUnpackA(self.context, value, r, g, b)
    return  r[0], g[0], b[0]
end

return color