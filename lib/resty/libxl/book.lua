require "resty.libxl.types.handle"
require "resty.libxl.types.book"
local ffi     = require "ffi"
local ffi_gc  = ffi.gc
local ffi_str = ffi.string
local lib     = require "resty.libxl.library"
local date    = require "resty.libxl.date"
local color   = require "resty.libxl.color"
local sheets  = require "resty.libxl.sheets"
local formats = require "resty.libxl.formats"
local fonts   = require "resty.libxl.fonts"

local book = {}

function book:__index(n)
    if n == "version" then
        return lib.xlBookVersionA(self.context)
    elseif n == "biff_version" or n == "biffversion" then
        return lib.xlBookBiffVersionA(self.context)
    elseif n == "ref_r1c1" or n == "refr1c1" then
        return lib.xlBookRefR1C1A(self.context)
    elseif n == "rgb_mode" or n == "rgbmode" then
        return lib.xlBookRgbModeA(self.context)
    elseif n == "is_date1904" or n == "isdate1904" then
        return lib.xlBookRgbModeA(self.context)
    elseif n == "is_template" or n == "istemplate" then
        return lib.xlBookIsTemplateA(self.context) == 1
    elseif n == "error_message" or n == "errormessage" then
        return ffi_str(lib.xlBookErrorMessageA(self.context))
    elseif n == "picture_size" or n == "picturesize" then
        return lib.xlBookPictureSizeA(self.context)
    else
        return rawget(book, n)
    end
end

function book:__newindex(n, v)
    if n == "ref_r1c1" or n == "refr1c1" then
        lib.xlBookSetRefR1C1A(self.context, v)
    elseif n == "rgb_mode" or n == "rgbmode" then
        lib.xlBookSetRgbModeA(self.context, v)
    elseif n == "date1904" then
        return lib.xlBookSetDate1904A(self.context)
    elseif n == "template" then
        return lib.xlBookSetTemplateA(self.context)
    elseif n == "locale" then
        return lib.xlBookSetLocaleA(self.context)
    else
        rawset(book, n, v)
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

