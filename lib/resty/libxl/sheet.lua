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

--[[
APIs not implemented:
FormatHandle __cdecl xlSheetCellFormatA(SheetHandle handle, int row, int col);
        void __cdecl xlSheetSetCellFormatA(SheetHandle handle, int row, int col, FormatHandle format);

         int __cdecl xlSheetWriteFormulaA(SheetHandle handle, int row, int col, const char* value, FormatHandle format);

 const char* __cdecl xlSheetReadCommentA(SheetHandle handle, int row, int col);
        void __cdecl xlSheetWriteCommentA(SheetHandle handle, int row, int col, const char* value, const char* author, int width, int height);

         int __cdecl xlSheetIsDateA(SheetHandle handle, int row, int col);
         int __cdecl xlSheetReadErrorA(SheetHandle handle, int row, int col);

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

         int __cdecl xlSheetPictureSizeA(SheetHandle handle);
         int __cdecl xlSheetGetPictureA(SheetHandle handle, int index, int* rowTop, int* colLeft, int* rowBottom, int* colRight,
                                                                                  int* width, int* height, int* offset_x, int* offset_y);

        void __cdecl xlSheetSetPictureA(SheetHandle handle, int row, int col, int pictureId, double scale, int offset_x, int offset_y);
        void __cdecl xlSheetSetPicture2A(SheetHandle handle, int row, int col, int pictureId, int width, int height, int offset_x, int offset_y);

         int __cdecl xlSheetGetHorPageBreakA(SheetHandle handle, int index);
         int __cdecl xlSheetGetHorPageBreakSizeA(SheetHandle handle);

         int __cdecl xlSheetGetVerPageBreakA(SheetHandle handle, int index);
         int __cdecl xlSheetGetVerPageBreakSizeA(SheetHandle handle);

         int __cdecl xlSheetSetHorPageBreakA(SheetHandle handle, int row, int pageBreak);
         int __cdecl xlSheetSetVerPageBreakA(SheetHandle handle, int col, int pageBreak);

        void __cdecl xlSheetSplitA(SheetHandle handle, int row, int col);
         int __cdecl xlSheetSplitInfoA(SheetHandle handle, int* row, int* col);

         int __cdecl xlSheetGroupRowsA(SheetHandle handle, int rowFirst, int rowLast, int collapsed);
         int __cdecl xlSheetGroupColsA(SheetHandle handle, int colFirst, int colLast, int collapsed);

         int __cdecl xlSheetGroupSummaryBelowA(SheetHandle handle);
        void __cdecl xlSheetSetGroupSummaryBelowA(SheetHandle handle, int below);

         int __cdecl xlSheetGroupSummaryRightA(SheetHandle handle);
        void __cdecl xlSheetSetGroupSummaryRightA(SheetHandle handle, int right);

         int __cdecl xlSheetInsertColA(SheetHandle handle, int colFirst, int colLast);
         int __cdecl xlSheetInsertRowA(SheetHandle handle, int rowFirst, int rowLast);
         int __cdecl xlSheetRemoveColA(SheetHandle handle, int colFirst, int colLast);
         int __cdecl xlSheetRemoveRowA(SheetHandle handle, int rowFirst, int rowLast);

         int __cdecl xlSheetCopyCellA(SheetHandle handle, int rowSrc, int colSrc, int rowDst, int colDst);

         int __cdecl xlSheetFirstRowA(SheetHandle handle);
         int __cdecl xlSheetLastRowA(SheetHandle handle);
         int __cdecl xlSheetFirstColA(SheetHandle handle);
         int __cdecl xlSheetLastColA(SheetHandle handle);

         int __cdecl xlSheetDisplayGridlinesA(SheetHandle handle);
        void __cdecl xlSheetSetDisplayGridlinesA(SheetHandle handle, int show);

         int __cdecl xlSheetPrintGridlinesA(SheetHandle handle);
        void __cdecl xlSheetSetPrintGridlinesA(SheetHandle handle, int print);

         int __cdecl xlSheetZoomA(SheetHandle handle);
        void __cdecl xlSheetSetZoomA(SheetHandle handle, int zoom);

         int __cdecl xlSheetPrintZoomA(SheetHandle handle);
        void __cdecl xlSheetSetPrintZoomA(SheetHandle handle, int zoom);

         int __cdecl xlSheetGetPrintFitA(SheetHandle handle, int* wPages, int* hPages);
        void __cdecl xlSheetSetPrintFitA(SheetHandle handle, int wPages, int hPages);

         int __cdecl xlSheetLandscapeA(SheetHandle handle);
        void __cdecl xlSheetSetLandscapeA(SheetHandle handle, int landscape);

         int __cdecl xlSheetPaperA(SheetHandle handle);
        void __cdecl xlSheetSetPaperA(SheetHandle handle, int paper);

 const char* __cdecl xlSheetHeaderA(SheetHandle handle);
         int __cdecl xlSheetSetHeaderA(SheetHandle handle, const char* header, double margin);
      double __cdecl xlSheetHeaderMarginA(SheetHandle handle);

 const char* __cdecl xlSheetFooterA(SheetHandle handle);
         int __cdecl xlSheetSetFooterA(SheetHandle handle, const char* footer, double margin);
      double __cdecl xlSheetFooterMarginA(SheetHandle handle);

         int __cdecl xlSheetHCenterA(SheetHandle handle);
        void __cdecl xlSheetSetHCenterA(SheetHandle handle, int hCenter);

         int __cdecl xlSheetVCenterA(SheetHandle handle);
        void __cdecl xlSheetSetVCenterA(SheetHandle handle, int vCenter);

      double __cdecl xlSheetMarginLeftA(SheetHandle handle);
        void __cdecl xlSheetSetMarginLeftA(SheetHandle handle, double margin);

      double __cdecl xlSheetMarginRightA(SheetHandle handle);
        void __cdecl xlSheetSetMarginRightA(SheetHandle handle, double margin);

      double __cdecl xlSheetMarginTopA(SheetHandle handle);
        void __cdecl xlSheetSetMarginTopA(SheetHandle handle, double margin);

      double __cdecl xlSheetMarginBottomA(SheetHandle handle);
        void __cdecl xlSheetSetMarginBottomA(SheetHandle handle, double margin);

         int __cdecl xlSheetPrintRowColA(SheetHandle handle);
        void __cdecl xlSheetSetPrintRowColA(SheetHandle handle, int print);

         int __cdecl xlSheetPrintRepeatRowsA(SheetHandle handle, int* rowFirst, int* rowLast);
        void __cdecl xlSheetSetPrintRepeatRowsA(SheetHandle handle, int rowFirst, int rowLast);

         int __cdecl xlSheetPrintRepeatColsA(SheetHandle handle, int* colFirst, int* colLast);
        void __cdecl xlSheetSetPrintRepeatColsA(SheetHandle handle, int colFirst, int colLast);

         int __cdecl xlSheetPrintAreaA(SheetHandle handle, int* rowFirst, int* rowLast, int* colFirst, int* colLast);
        void __cdecl xlSheetSetPrintAreaA(SheetHandle handle, int rowFirst, int rowLast, int colFirst, int colLast);

        void __cdecl xlSheetClearPrintRepeatsA(SheetHandle handle);
        void __cdecl xlSheetClearPrintAreaA(SheetHandle handle);

         int __cdecl xlSheetGetNamedRangeA(SheetHandle handle, const char* name, int* rowFirst, int* rowLast, int* colFirst, int* colLast, int scopeId, int* hidden);
         int __cdecl xlSheetSetNamedRangeA(SheetHandle handle, const char* name, int rowFirst, int rowLast, int colFirst, int colLast, int scopeId);
         int __cdecl xlSheetDelNamedRangeA(SheetHandle handle, const char* name, int scopeId);

         int __cdecl xlSheetNamedRangeSizeA(SheetHandle handle);
 const char* __cdecl xlSheetNamedRangeA(SheetHandle handle, int index, int* rowFirst, int* rowLast, int* colFirst, int* colLast, int* scopeId, int* hidden);

         int __cdecl xlSheetHyperlinkSizeA(SheetHandle handle);
 const char* __cdecl xlSheetHyperlinkA(SheetHandle handle, int index, int* rowFirst, int* rowLast, int* colFirst, int* colLast);
         int __cdecl xlSheetDelHyperlinkA(SheetHandle handle, int index);
        void __cdecl xlSheetAddHyperlinkA(SheetHandle handle, const char* hyperlink, int rowFirst, int rowLast, int colFirst, int colLast);

        void __cdecl xlSheetGetTopLeftViewA(SheetHandle handle, int* row, int* col);
        void __cdecl xlSheetSetTopLeftViewA(SheetHandle handle, int row, int col);

         int __cdecl xlSheetRightToLeftA(SheetHandle handle);
        void __cdecl xlSheetSetRightToLeftA(SheetHandle handle, int rightToLeft);

        void __cdecl xlSheetSetAutoFitAreaA(SheetHandle handle, int rowFirst, int colFirst, int rowLast, int colLast);

        void __cdecl xlSheetAddrToRowColA(SheetHandle handle, const char* addr, int* row, int* col, int* rowRelative, int* colRelative);
 const char* __cdecl xlSheetRowColToAddrA(SheetHandle handle, int row, int col, int rowRelative, int colRelative);

 ]]

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

function sheet:clear(rf, rl, cf, cl)
    rf = rf or 0
    rl = rl or 1048575
    cf = cf or 0
    cl = cl or 16383
    if rf < 0 then rf = 0 end
    if rf > 1048575 then rf = 1048575 end
    if rl < 0 then rl = 0 end
    if rl > 1048575 then rl = 1048575 end
    if cf < 0 then cf = 0 end
    if cf > 16383 then cf = 16383 end
    if cl < 0 then cl = 0 end
    if cl > 16383 then cl = 16383 end
    lib.xlSheetClearA(self.context, rf, rl, cf, cl)
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
    elseif n == "protect" then
        return lib.xlSheetProtectA(self.context) == 1
    elseif n == "hidden" then
        return lib.xlSheetHiddenA(self.context) == 1
    else
        return rawget(sheet, n)
    end
end

function sheet:__newindex(n, v)
    if n == "name" then
        lib.xlSheetSetNameA(self.context, v)
    elseif n == "protect" then
        lib.xlSheetSetProtectA(self.context, v)
    elseif n == "hidden" then
        lib.xlSheetSetHiddenA(self.context, v)
    else
        rawset(self, n, v)
    end
end

return sheet
