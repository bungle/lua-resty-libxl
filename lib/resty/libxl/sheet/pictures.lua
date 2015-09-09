local lib          = require "resty.libxl.library"
local setmetatable = setmetatable
local rawget       = rawget
local type         = type
local ffi          = require "ffi"
local ffi_new      = ffi.new
local pictures     = {}

local rt = ffi_new("int[1]", 0)
local cl = ffi_new("int[1]", 0)
local rb = ffi_new("int[1]", 0)
local cr = ffi_new("int[1]", 0)
local w = ffi_new("int[1]", 0)
local h = ffi_new("int[1]", 0)
local ox = ffi_new("int[1]", 0)
local oy = ffi_new("int[1]", 0)

function pictures.new(opts)
    return setmetatable(opts, pictures)
end

function pictures:set(row, col, picture)
    row, col = row - 1, col -1
    local t = type(picture)
    if t == "number" then
        lib.xlSheetSetPictureA(self.sheet.context, row, col, picture - 1, 1.0, 0, 0)
    elseif t == "table" then
        local i = picture.index - 1
        local x, y = 0, 0
        local t = type(picture.offset)
        if t == "table" then
            x = picture.offset.x or x
            y = picture.offset.y or y
        else
            x = picture.offset_x or x
            y = picture.offset_y or y
        end
        if picture.width or picture.height then
            local w = picture.width or -1
            local h = picture.height or -1
            lib.xlSheetSetPicture2A(self.sheet.context, row, col, i, w, h, x, y)
        else
            local s = picture.scale or 1.0
            lib.xlSheetSetPictureA(self.sheet.context, row, col, i, s, x, y)
        end
    end
end

function pictures:__len()
    return lib.xlSheetPictureSizeA(self.sheet.context)
end

function pictures:__index(n)
    if n == "size" or n == "count" or n == "n" then
        return lib.xlSheetPictureSizeA(self.sheet.context)
    elseif type(n) == "number" then
        local i = lib.xlSheetGetPictureA(self.sheet.context, n - 1, rt, cl, rb, cr, w, h, ox, oy)
        if i == -1 then
            return nil, self.sheet.book.error
        else
            return {
                index = i + 1,
                row = {
                    top = rt,
                    bottom = rb
                }, col = {
                    left = cl,
                    right = cr
                }, width = w,
                height = h,
                offset = {
                    x = ox,
                    y = oy
                }
            }
        end
    else
        return rawget(pictures, n)
    end
end

return pictures