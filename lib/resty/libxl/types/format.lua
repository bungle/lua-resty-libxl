local ffi      = require "ffi"
local ffi_cdef = ffi.cdef

ffi_cdef[[
FontHandle __cdecl xlFormatFontA(FormatHandle handle);
       int __cdecl xlFormatSetFontA(FormatHandle handle, FontHandle fontHandle);

       int __cdecl xlFormatNumFormatA(FormatHandle handle);
      void __cdecl xlFormatSetNumFormatA(FormatHandle handle, int numFormat);

       int __cdecl xlFormatAlignHA(FormatHandle handle);
      void __cdecl xlFormatSetAlignHA(FormatHandle handle, int align);

       int __cdecl xlFormatAlignVA(FormatHandle handle);
      void __cdecl xlFormatSetAlignVA(FormatHandle handle, int align);

       int __cdecl xlFormatWrapA(FormatHandle handle);
      void __cdecl xlFormatSetWrapA(FormatHandle handle, int wrap);

       int __cdecl xlFormatRotationA(FormatHandle handle);
       int __cdecl xlFormatSetRotationA(FormatHandle handle, int rotation);

       int __cdecl xlFormatIndentA(FormatHandle handle);
      void __cdecl xlFormatSetIndentA(FormatHandle handle, int indent);

       int __cdecl xlFormatShrinkToFitA(FormatHandle handle);
      void __cdecl xlFormatSetShrinkToFitA(FormatHandle handle, int shrinkToFit);

      void __cdecl xlFormatSetBorderA(FormatHandle handle, int style);
      void __cdecl xlFormatSetBorderColorA(FormatHandle handle, int color);

       int __cdecl xlFormatBorderLeftA(FormatHandle handle);
      void __cdecl xlFormatSetBorderLeftA(FormatHandle handle, int style);

       int __cdecl xlFormatBorderRightA(FormatHandle handle);
      void __cdecl xlFormatSetBorderRightA(FormatHandle handle, int style);

       int __cdecl xlFormatBorderTopA(FormatHandle handle);
      void __cdecl xlFormatSetBorderTopA(FormatHandle handle, int style);

       int __cdecl xlFormatBorderBottomA(FormatHandle handle);
      void __cdecl xlFormatSetBorderBottomA(FormatHandle handle, int style);

       int __cdecl xlFormatBorderLeftColorA(FormatHandle handle);
      void __cdecl xlFormatSetBorderLeftColorA(FormatHandle handle, int color);

       int __cdecl xlFormatBorderRightColorA(FormatHandle handle);
      void __cdecl xlFormatSetBorderRightColorA(FormatHandle handle, int color);

       int __cdecl xlFormatBorderTopColorA(FormatHandle handle);
      void __cdecl xlFormatSetBorderTopColorA(FormatHandle handle, int color);

       int __cdecl xlFormatBorderBottomColorA(FormatHandle handle);
      void __cdecl xlFormatSetBorderBottomColorA(FormatHandle handle, int color);

       int __cdecl xlFormatBorderDiagonalA(FormatHandle handle);
      void __cdecl xlFormatSetBorderDiagonalA(FormatHandle handle, int border);

       int __cdecl xlFormatBorderDiagonalStyleA(FormatHandle handle);
      void __cdecl xlFormatSetBorderDiagonalStyleA(FormatHandle handle, int style);

       int __cdecl xlFormatBorderDiagonalColorA(FormatHandle handle);
      void __cdecl xlFormatSetBorderDiagonalColorA(FormatHandle handle, int color);

       int __cdecl xlFormatFillPatternA(FormatHandle handle);
      void __cdecl xlFormatSetFillPatternA(FormatHandle handle, int pattern);

       int __cdecl xlFormatPatternForegroundColorA(FormatHandle handle);
      void __cdecl xlFormatSetPatternForegroundColorA(FormatHandle handle, int color);

       int __cdecl xlFormatPatternBackgroundColorA(FormatHandle handle);
      void __cdecl xlFormatSetPatternBackgroundColorA(FormatHandle handle, int color);

       int __cdecl xlFormatLockedA(FormatHandle handle);
      void __cdecl xlFormatSetLockedA(FormatHandle handle, int locked);

       int __cdecl xlFormatHiddenA(FormatHandle handle);
      void __cdecl xlFormatSetHiddenA(FormatHandle handle, int hidden);
]]