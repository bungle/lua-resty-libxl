require "resty.libxl.types.handle"
require "resty.libxl.types.sheet"
local setmetatable = setmetatable
local rawget       = rawget
local rawset       = rawset
local type         = type
local ffi          = require "ffi"
local C            = ffi.C
local ffi_str      = ffi.string
local lib          = require "resty.libxl.library"
local pictures     = require "resty.libxl.sheet.pictures"
local hyperlinks   = require "resty.libxl.sheet.hyperlinks"

--[[
APIs not implemented:
FormatHandle __cdecl xlSheetCellFormatA(SheetHandle handle, int row, int col);
        void __cdecl xlSheetSetCellFormatA(SheetHandle handle, int row, int col, FormatHandle format);

 const char* __cdecl xlSheetReadCommentA(SheetHandle handle, int row, int col);
        void __cdecl xlSheetWriteCommentA(SheetHandle handle, int row, int col, const char* value, const char* author, int width, int height);

      double __cdecl xlSheetColWidthA(SheetHandle handle, int col);
      double __cdecl xlSheetRowHeightA(SheetHandle handle, int row);

         int __cdecl xlSheetSetColA(SheetHandle handle, int colFirst, int colLast, double width, FormatHandle format, int hidden);
         int __cdecl xlSheetSetRowA(SheetHandle handle, int row, double height, FormatHandle format, int hidden);

         int __cdecl xlSheetRowHiddenA(SheetHandle handle, int row);
         int __cdecl xlSheetSetRowHiddenA(SheetHandle handle, int row, int hidden);

         int __cdecl xlSheetColHiddenA(SheetHandle handle, int col);
         int __cdecl xlSheetSetColHiddenA(SheetHandle handle, int col, int hidden);

         int __cdecl xlSheetGetMergeA(SheetHandle handle, int row, int col, int* rowFirst, int* rowLast, int* colFirst, int* colLast);
         int __cdecl xlSheetSetMergeA(SheetHandle handle, int rowFirst, int rowLast, int colFirst, int colLast);
         int __cdecl xlSheetDelMergeA(SheetHandle handle, int row, int col);

         int __cdecl xlSheetMergeSizeA(SheetHandle handle);
         int __cdecl xlSheetMergeA(SheetHandle handle, int index, int* rowFirst, int* rowLast, int* colFirst, int* colLast);
         int __cdecl xlSheetDelMergeByIndexA(SheetHandle handle, int index);

         int __cdecl xlSheetGetHorPageBreakA(SheetHandle handle, int index);
         int __cdecl xlSheetGetHorPageBreakSizeA(SheetHandle handle);

         int __cdecl xlSheetGetVerPageBreakA(SheetHandle handle, int index);
         int __cdecl xlSheetGetVerPageBreakSizeA(SheetHandle handle);

        void __cdecl xlSheetSplitA(SheetHandle handle, int row, int col);
         int __cdecl xlSheetSplitInfoA(SheetHandle handle, int* row, int* col);

         int __cdecl xlSheetGroupRowsA(SheetHandle handle, int rowFirst, int rowLast, int collapsed);
         int __cdecl xlSheetGroupColsA(SheetHandle handle, int colFirst, int colLast, int collapsed);

         int __cdecl xlSheetInsertColA(SheetHandle handle, int colFirst, int colLast);
         int __cdecl xlSheetInsertRowA(SheetHandle handle, int rowFirst, int rowLast);
         int __cdecl xlSheetRemoveColA(SheetHandle handle, int colFirst, int colLast);
         int __cdecl xlSheetRemoveRowA(SheetHandle handle, int rowFirst, int rowLast);

         int __cdecl xlSheetCopyCellA(SheetHandle handle, int rowSrc, int colSrc, int rowDst, int colDst);

         int __cdecl xlSheetGetPrintFitA(SheetHandle handle, int* wPages, int* hPages);
        void __cdecl xlSheetSetPrintFitA(SheetHandle handle, int wPages, int hPages);

         int __cdecl xlSheetSetHeaderA(SheetHandle handle, const char* header, double margin);
         int __cdecl xlSheetSetFooterA(SheetHandle handle, const char* footer, double margin);

         int __cdecl xlSheetPrintRepeatRowsA(SheetHandle handle, int* rowFirst, int* rowLast);
        void __cdecl xlSheetSetPrintRepeatRowsA(SheetHandle handle, int rowFirst, int rowLast);

         int __cdecl xlSheetPrintRepeatColsA(SheetHandle handle, int* colFirst, int* colLast);
        void __cdecl xlSheetSetPrintRepeatColsA(SheetHandle handle, int colFirst, int colLast);

         int __cdecl xlSheetPrintAreaA(SheetHandle handle, int* rowFirst, int* rowLast, int* colFirst, int* colLast);
        void __cdecl xlSheetSetPrintAreaA(SheetHandle handle, int rowFirst, int rowLast, int colFirst, int colLast);

         int __cdecl xlSheetGetNamedRangeA(SheetHandle handle, const char* name, int* rowFirst, int* rowLast, int* colFirst, int* colLast, int scopeId, int* hidden);
         int __cdecl xlSheetSetNamedRangeA(SheetHandle handle, const char* name, int rowFirst, int rowLast, int colFirst, int colLast, int scopeId);
         int __cdecl xlSheetDelNamedRangeA(SheetHandle handle, const char* name, int scopeId);

         int __cdecl xlSheetNamedRangeSizeA(SheetHandle handle);
 const char* __cdecl xlSheetNamedRangeA(SheetHandle handle, int index, int* rowFirst, int* rowLast, int* colFirst, int* colLast, int* scopeId, int* hidden);

        void __cdecl xlSheetGetTopLeftViewA(SheetHandle handle, int* row, int* col);
        void __cdecl xlSheetSetTopLeftViewA(SheetHandle handle, int row, int col);

        void __cdecl xlSheetAddrToRowColA(SheetHandle handle, const char* addr, int* row, int* col, int* rowRelative, int* colRelative);
 const char* __cdecl xlSheetRowColToAddrA(SheetHandle handle, int row, int col, int rowRelative, int colRelative);
 ]]

local sheet = {}

function sheet.new(opts)
    local self = setmetatable(opts, sheet)
    self.pictures = pictures.new{ sheet = self }
    self.hyperlinks = hyperlinks.new{ sheet = self }
    return self
end

function sheet:write(row, col, value, format)
    row, col = row - 1, col - 1
    local  r
    local  t = type(value)
    if     t == "string" then
        r = lib.xlSheetWriteStrA(self.context, row, col, value, format or 0)
    elseif t == "number" then
        r = lib.xlSheetWriteNumA(self.context, row, col, value, format or 0)
    elseif t == "boolean" then
        r = lib.xlSheetWriteBoolA(self.context, row, col, value, format or 0)
    elseif t == "nil" then
        r = lib.xlSheetWriteBlankA(self.context, row, col, format or 0)
    elseif t == "table" then
        r = lib.xlSheetWriteStrA(self.context, row, col, tostring(value), format or 0)
    end
    if r == 0 then
        return nil, self.book.error
    end
    return true
end

function sheet:writeformula(row, col, value, format)
    if lib.xlSheetWriteFormulaA(self.context, row - 1, col - 1, value, format) == 0 then
        return nil, self.book.error
    end
    return true
end

function sheet:writedate(row, col, year, month, day, hour, min, sec, msec, format)
    local date = self.book.date:pack(year, month, day, hour, min, sec, msec)
    return self:write(row, col, date, format)
end

function sheet:read(row, col, format)
    local r, c = row - 1, col - 1
    local  t = self:celltype(row, col)
    if     t == C.CELLTYPE_EMPTY   then
        return ""
    elseif t == C.CELLTYPE_NUMBER  then
        return lib.xlSheetReadNumA(self.context, r, c, format)
    elseif t == C.CELLTYPE_STRING  then
        return ffi_str(lib.xlSheetReadStrA(self.context, r, c, format))
    elseif t == C.CELLTYPE_BOOLEAN then
        return lib.xlSheetReadBoolA(self.context, r, c, format) ~= 0
    elseif t == C.CELLTYPE_BLANK   then
        lib.xlSheetReadBlankA(self.context, r, c, format)
        return nil
    elseif t == C.CELLTYPE_ERROR   then
        return lib.xlSheetReadErrorA(self.context, r, c)
    else
        return nil
    end
end

function sheet:readformula(row, col, format)
    if self:isformula(row, col) then
        local formula = lib.xlSheetReadFormulaA(self.context, row - 1, col - 1, format)
        if formula == nil then
            return nil, self.book.error
        end
        return ffi_str(formula)
    end
    return nil
end

function sheet:readdate(row, col, format)
    if self:isdate(row, col) then
        return self.book.date:unpack(self:read(row, col, format))
    end
    return nil
end

function sheet:clear(rf, rl, cf, cl)
    rf = rf or 1
    rl = rl or 1048576
    cf = cf or 1
    cl = cl or 16384
    rf, rl, cf, cl = rf - 1, rl - 1, cf - 1, cl - 1
    if rf < 0 then
        rf = 0
    elseif rf > 1048575 then
        rf = 1048575
    end
    if rl < 0 then
        rl = 0
    elseif rl > 1048575 then
        rl = 1048575
    end
    if cf < 0 then
        cf = 0
    elseif cf > 16383 then
        cf = 16383
    end
    if cl < 0 then
        cl = 0
    elseif cl > 16383 then
        cl = 16383
    end
    lib.xlSheetClearA(self.context, rf, rl, cf, cl)
end

function sheet:clearprint(type)
    type = type or "area"
    if type == "all" then
        lib.xlSheetClearPrintAreaA(self.context)
        lib.xlSheetClearPrintRepeatsA(self.context)
    elseif type == "area" then
        lib.xlSheetClearPrintAreaA(self.context)
    elseif type == "repeats" then
        lib.xlSheetClearPrintRepeatsA(self.context)
    end
end

function sheet:autofit(rf, cf, rl, cl)
    lib.xlSheetSetAutoFitAreaA(self.context, rf - 1, cf - 1, rl - 1, cl - 1)
end

function sheet:isformula(row, col)
    return lib.xlSheetIsFormulaA(self.context, row - 1, col - 1) == 1
end

function sheet:isdate(row, col)
    print(lib.xlSheetIsDateA(self.context, row - 1, col - 1))
    return lib.xlSheetIsDateA(self.context, row - 1, col - 1) == 1
end

function sheet:celltype(row, col)
    return lib.xlSheetCellTypeA(self.context, row - 1, col - 1)
end

function sheet:__index(n)
    if n == "name" then
        return ffi_str(lib.xlSheetNameA(self.context))
    elseif n == "landscape" then
        return lib.xlSheetLandscapeA(self.context) == 1
    elseif n == "zoom" then
        return lib.xlSheetZoomA(self.context)
    elseif n == "protect" then
        return lib.xlSheetProtectA(self.context) == 1
    elseif n == "hidden" then
        return lib.xlSheetHiddenA(self.context) == 1
    elseif n == "first-row" then
        return lib.xlSheetFirstRowA(self.context) + 1
    elseif n == "first-col" then
        return lib.xlSheetFirstColA(self.context) + 1
    elseif n == "last-row" then
        return lib.xlSheetLastRowA(self.context) + 1
    elseif n == "last-col" then
        return lib.xlSheetLastColA(self.context) + 1
    elseif n == "group-summary-below" then
        return lib.xlSheetGroupSummaryBelowA(self.context) == 1
    elseif n == "group-summary-right" then
        return lib.xlSheetGroupSummaryRightA(self.context) == 1
    elseif n == "display-gridlines" then
        return lib.xlSheetDisplayGridlinesA(self.context) == 1
    elseif n == "print-gridlines" then
        return lib.xlSheetPrintGridlinesA(self.context) == 1
    elseif n == "print-row-col" then
        return lib.xlSheetPrintRowColA(self.context) == 1
    elseif n == "print-zoom" then
        return lib.xlSheetPrintZoomA(self.context)
    elseif n == "header" then
        return lib.xlSheetHeaderA(self.context)
    elseif n == "header-margin" then
        return lib.xlSheetHeaderMarginA(self.context)
    elseif n == "footer" then
        return lib.xlSheetFooterA(self.context)
    elseif n == "footer-margin" then
        return lib.xlSheetFooterMarginA(self.context)
    elseif n == "margin-left" then
        return lib.xlSheetMarginLeftA(self.context)
    elseif n == "margin-right" then
        return lib.xlSheetMarginRightA(self.context)
    elseif n == "margin-top" then
        return lib.xlSheetMarginTopA(self.context)
    elseif n == "margin-bottom" then
        return lib.xlSheetMarginBottomA(self.context)
    elseif n == "paper" then
        return lib.xlSheetPaperA(self.context)
    elseif n == "h-center" then
        return lib.xlSheetHCenterA(self.context) == 1
    elseif n == "v-center" then
        return lib.xlSheetVCenterA(self.context) == 1
    elseif n == "right-to-left" then
        return lib.xlSheetRightToLeftA(self.context) == 1
    else
        return rawget(sheet, n)
    end
end

function sheet:__newindex(n, v)
    if n == "name" then
        lib.xlSheetSetNameA(self.context, v)
    elseif n == "landscape" then
        lib.xlSheetSetLandscapeA(self.context, v)
    elseif n == "zoom" then
        lib.xlSheetSetZoomA(self.context, v)
    elseif n == "protect" then
        lib.xlSheetSetProtectA(self.context, v)
    elseif n == "hidden" then
        lib.xlSheetSetHiddenA(self.context, v)
    elseif n == "group-summary-below" then
        lib.xlSheetSetGroupSummaryBelowA(self.context, v)
    elseif n == "group-summary-right" then
        lib.xlSheetSetGroupSummaryRightA(self.context, v)
    elseif n == "display-gridlines" then
        lib.xlSheetSetDisplayGridlinesA(self.context, v)
    elseif n == "print-gridlines" then
        lib.xlSheetSetPrintGridlinesA(self.context, v)
    elseif n == "print-row-col" then
        lib.xlSheetSetPrintRowColA(self.context, v)
    elseif n == "print-zoom" then
        lib.xlSheetSetPrintZoomA(self.context, v)
    elseif n == "margin-left" then
        lib.xlSheetSetMarginLeftA(self.context, v)
    elseif n == "margin-right" then
        lib.xlSheetSetMarginRightA(self.context, v)
    elseif n == "margin-top" then
        lib.xlSheetSetMarginTopA(self.context, v)
    elseif n == "margin-bottom" then
        lib.xlSheetSetMarginBottomA(self.context, v)
    elseif n == "paper" then
        lib.xlSheetSetPaperA(self.context, v)
    elseif n == "h-center" then
        lib.xlSheetSetHCenterA(self.context, v)
    elseif n == "v-center" then
        lib.xlSheetSetVCenterA(self.context, v)
    elseif n == "right-to-left" then
        lib.xlSheetSetRightToLeftA(self.context, v)
    else
        rawset(self, n, v)
    end
end

return sheet
