local lib          = require "resty.libxl.library"
local setmetatable = setmetatable
local ffi          = require "ffi"
local ffi_new      = ffi.new

local year  = ffi_new("int[1]", 0)
local month = ffi_new("int[1]", 0)
local day   = ffi_new("int[1]", 0)
local hour  = ffi_new("int[1]", 0)
local min   = ffi_new("int[1]", 0)
local sec   = ffi_new("int[1]", 0)
local msec  = ffi_new("int[1]", 0)

local date   = {}
date.__index = date

function date.new(opts)
    return setmetatable(opts, date)
end

function date:pack(year, month, day, hour, min, sec, msec)
    return lib.xlBookDatePackA(self.context, year or 0, month or 1, day or 1, hour or 0, min or 0, sec or 0, msec or 0)
end

function date:unpack(value)
    lib.xlBookDateUnpackA(self.context, value, year, month, day, hour, min, sec, msec)
    return  year[0], month[0], day[0], hour[0], min[0], sec[0], msec[0]
end

return date