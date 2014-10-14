require "resty.libxl.types.handle"
require "resty.libxl.types.book"
local ffi     = require "ffi"
local ffi_gc  = ffi.gc
local lib     = require "resty.libxl.library"
local date    = require "resty.libxl.date"
local color   = require "resty.libxl.color"
local sheets  = require "resty.libxl.sheets"
local formats = require "resty.libxl.formats"
local fonts   = require "resty.libxl.fonts"

local book = {}
book.__index = book

function book.new(opts)
    opts = opts or {}
    local context = opts.format == "xls" and ffi_gc(lib.xlCreateBookCA(), lib.xlBookReleaseA) or  ffi_gc(lib.xlCreateXMLBookCA(), lib.xlBookReleaseA)
    if type(opts.name) == "string" and type(opts.key) == "string" then
        lib.xlBookSetKeyA(context, opts.name, opts.key)
    end
    local self = setmetatable({
        context = context,
           date = date.new{ context = context },
          color = color.new{ context = context },
    }, book)
    book.sheets = sheets.new{ book = self }
    book.formats = formats.new{ book = self }
    book.fonts = fonts.new{ book = self }
    return self
end

function book:load(filename)
    return lib.xlBookLoadA(self.context, filename) == 1
end

function book:save(filename)
    return lib.xlBookSaveA(self.context, filename) == 1
end

function book:release()
    return lib.xlBookReleaseA(self.context)
end

return book

