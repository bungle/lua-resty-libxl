local ffi        = require "ffi"
local ffi_cdef   = ffi.cdef

ffi_cdef[[
  BookHandle __cdecl xlCreateBookCA(void);
  BookHandle __cdecl xlCreateXMLBookCA(void);

         int __cdecl xlBookLoadA(BookHandle handle, const char* filename);
         int __cdecl xlBookSaveA(BookHandle handle, const char* filename);

         int __cdecl xlBookLoadRawA(BookHandle handle, const char* data, unsigned size);
         int __cdecl xlBookSaveRawA(BookHandle handle, const char** data, unsigned* size);

 SheetHandle __cdecl xlBookAddSheetA(BookHandle handle, const char* name, SheetHandle initSheet);
 SheetHandle __cdecl xlBookInsertSheetA(BookHandle handle, int index, const char* name, SheetHandle initSheet);
 SheetHandle __cdecl xlBookGetSheetA(BookHandle handle, int index);
         int __cdecl xlBookSheetTypeA(BookHandle handle, int index);
         int __cdecl xlBookDelSheetA(BookHandle handle, int index);
         int __cdecl xlBookSheetCountA(BookHandle handle);

FormatHandle __cdecl xlBookAddFormatA(BookHandle handle, FormatHandle initFormat);
  FontHandle __cdecl xlBookAddFontA(BookHandle handle, FontHandle initFont);
         int __cdecl xlBookAddCustomNumFormatA(BookHandle handle, const char* customNumFormat);
 const char* __cdecl xlBookCustomNumFormatA(BookHandle handle, int fmt);

FormatHandle __cdecl xlBookFormatA(BookHandle handle, int index);
         int __cdecl xlBookFormatSizeA(BookHandle handle);

  FontHandle __cdecl xlBookFontA(BookHandle handle, int index);
         int __cdecl xlBookFontSizeA(BookHandle handle);

      double __cdecl xlBookDatePackA(BookHandle handle, int year, int month, int day, int hour, int min, int sec, int msec);
         int __cdecl xlBookDateUnpackA(BookHandle handle, double value, int* year, int* month, int* day, int* hour, int* min, int* sec, int* msec);

         int __cdecl xlBookColorPackA(BookHandle handle, int red, int green, int blue);
        void __cdecl xlBookColorUnpackA(BookHandle handle, int color, int* red, int* green, int* blue);

         int __cdecl xlBookActiveSheetA(BookHandle handle);
        void __cdecl xlBookSetActiveSheetA(BookHandle handle, int index);

         int __cdecl xlBookPictureSizeA(BookHandle handle);
         int __cdecl xlBookGetPictureA(BookHandle handle, int index, const char** data, unsigned* size);

         int __cdecl xlBookAddPictureA(BookHandle handle, const char* filename);
         int __cdecl xlBookAddPicture2A(BookHandle handle, const char* data, unsigned size);

 const char* __cdecl xlBookDefaultFontA(BookHandle handle, int* fontSize);
        void __cdecl xlBookSetDefaultFontA(BookHandle handle, const char* fontName, int fontSize);

         int __cdecl xlBookRefR1C1A(BookHandle handle);
        void __cdecl xlBookSetRefR1C1A(BookHandle handle, int refR1C1);

        void __cdecl xlBookSetKeyA(BookHandle handle, const char* name, const char* key);

         int __cdecl xlBookRgbModeA(BookHandle handle);
        void __cdecl xlBookSetRgbModeA(BookHandle handle, int rgbMode);

         int __cdecl xlBookVersionA(BookHandle handle);
         int __cdecl xlBookBiffVersionA(BookHandle handle);

         int __cdecl xlBookIsDate1904A(BookHandle handle);
        void __cdecl xlBookSetDate1904A(BookHandle handle, int date1904);

         int __cdecl xlBookIsTemplateA(BookHandle handle);
        void __cdecl xlBookSetTemplateA(BookHandle handle, int tmpl);

         int __cdecl xlBookSetLocaleA(BookHandle handle, const char* locale);
 const char* __cdecl xlBookErrorMessageA(BookHandle handle);

        void __cdecl xlBookReleaseA(BookHandle handle);
]]
