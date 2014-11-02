require "resty.libxl.types.handle"
require "resty.libxl.types.sheet"
local ffi     = require "ffi"
local C       = ffi.C
local ffi_str = ffi.string
local lib     = require "resty.libxl.library"

local sheet = {}

function sheet.new(opts)
    return setmetatable(opts, sheet)
end

function sheet:write(row, col, value, format)
    row, col = row - 1, col - 1
    local  t = type(value)
    if     t == "string" then
        lib.xlSheetWriteStrA(self.context, row, col, value, format)
    elseif t == "number" then
        lib.xlSheetWriteNumA(self.context, row, col, value, format)
    elseif t == "boolean" then
        lib.xlSheetWriteBoolA(self.context, row, col, value, format)
    elseif t == "nil" then
        lib.xlSheetWriteBlankA(self.context, row, col, format)
    elseif t == "table" then
        lib.xlSheetWriteStrA(self.context, row, col, tostring(value), format)
    end
    return self
end

function sheet:read(row, col, format)
    local row,col = row - 1,col - 1
    if self:isformula(row, col) then
        return ffi_str(lib.xlSheetReadFormulaA(self.context, row, col, format))
    elseif self:isdate(row, col) then
        return self.book.date:unpack(lib.xlSheetReadNumA(self.context, row, col, format))
    else
        local  t = self:celltype(row, col)
        if     t == C.CELLTYPE_EMPTY   then
            return ""
        elseif t == C.CELLTYPE_NUMBER  then
            return lib.xlSheetReadNumA(self.context, row, col, format)
        elseif t == C.CELLTYPE_STRING  then
            return ffi_str(lib.xlSheetReadStrA(self.context, row, col, format))
        elseif t == C.CELLTYPE_BOOLEAN then
            return lib.xlSheetReadBoolA(self.context, row, col, format) ~= 0
        elseif t == C.CELLTYPE_BLANK   then
            lib.xlSheetReadBlankA(self.context, row, col, format)
            return nil
        elseif t == C.CELLTYPE_ERROR   then
            return lib.xlSheetReadErrorA(self.context, row, col)
        else
            return nil
        end
    end
end

function sheet:isformula(row, col)
    return lib.xlSheetIsFormulaA(self.context, row - 1, col - 1) == 1
end

function sheet:isdate(row, col)
    return lib.xlSheetIsDateA(self.context, row - 1, col - 1) == 1
end

function sheet:celltype(row, col)
    return lib.xlSheetCellTypeA(self.context, row - 1, col - 1)
end

function sheet:__index(n)
    if n == "name" then
        return ffi_str(lib.xlSheetNameA(self.context))
    else
        rawget(sheet, n)
    end
end

function sheet:__newindex(n, v)
    if n == "name" then
        lib.xlSheetSetNameA(self.context, v)
    else
        rawset(sheet, n, v)
    end
end

return sheet
