local lib    = require "resty.libxl.library"
local sheet  = require "resty.libxl.sheet"
local sheets = {}

function sheets.new(opts)
    return setmetatable(opts, sheets)
end

function sheets:add(name, init)
    lib.xlBookAddSheetA(self.book.context, name, init and init.context or init)
    return self[self.size]
end

function sheets:insert(index, name, init)
    lib.xlBookInsertSheetA(self.book.context, index, name, init and init.context or init)
    return self[index-1]
end

function sheets:del(index)
    return lib.xlBookDelSheetA(self.book.context, index - 1) == 1
end

function sheets:type(index)
    return lib.xlBookSheetTypeA(self.book.context, index - 1)
end

function sheets:__len()
    return lib.xlBookSheetCountA(self.book.context)
end

function sheets:__index(n)
    if n == "active" then
        return lib.xlBookActiveSheetA(self.book.context) + 1
    elseif n == "size" or n == "count" then
        return lib.xlBookSheetCountA(self.book.context)
    elseif type(n) == "number" then
        return sheet.new{ context = lib.xlBookGetSheetA(self.book.context, n - 1), type = lib.xlBookSheetTypeA(self.book.context, n - 1), book = self.book }
    else
        return rawget(sheets, n)
    end
end

function sheets:__newindex(n, v)
    if n == "active" then
        lib.xlBookSetActiveSheetA(self.book.context, v - 1)
    else
        rawset(sheets, n, v)
    end
end

return sheets