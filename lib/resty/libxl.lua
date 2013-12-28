local ffi        = require "ffi"
local ffi_new    = ffi.new
local ffi_cdef   = ffi.cdef
local ffi_load   = ffi.load
local C          = ffi.C
local ffi_string = ffi.string

local libxl = ffi_load("libxl")

ffi_cdef[[
typedef enum {COLOR_BLACK = 8, COLOR_WHITE, COLOR_RED, COLOR_BRIGHTGREEN,
              COLOR_BLUE, COLOR_YELLOW, COLOR_PINK, COLOR_TURQUOISE,
              COLOR_DARKRED, COLOR_GREEN, COLOR_DARKBLUE, COLOR_DARKYELLOW,
              COLOR_VIOLET, COLOR_TEAL, COLOR_GRAY25, COLOR_GRAY50,
              COLOR_PERIWINKLE_CF, COLOR_PLUM_CF, COLOR_IVORY_CF,
              COLOR_LIGHTTURQUOISE_CF, COLOR_DARKPURPLE_CF, COLOR_CORAL_CF,
              COLOR_OCEANBLUE_CF, COLOR_ICEBLUE_CF, COLOR_DARKBLUE_CL,
              COLOR_PINK_CL, COLOR_YELLOW_CL, COLOR_TURQUOISE_CL,
              COLOR_VIOLET_CL, COLOR_DARKRED_CL, COLOR_TEAL_CL, COLOR_BLUE_CL,
              COLOR_SKYBLUE, COLOR_LIGHTTURQUOISE, COLOR_LIGHTGREEN,
              COLOR_LIGHTYELLOW, COLOR_PALEBLUE, COLOR_ROSE, COLOR_LAVENDER,
              COLOR_TAN, COLOR_LIGHTBLUE, COLOR_AQUA, COLOR_LIME, COLOR_GOLD,
              COLOR_LIGHTORANGE, COLOR_ORANGE, COLOR_BLUEGRAY, COLOR_GRAY40,
              COLOR_DARKTEAL, COLOR_SEAGREEN, COLOR_DARKGREEN,
              COLOR_OLIVEGREEN, COLOR_BROWN, COLOR_PLUM, COLOR_INDIGO,
              COLOR_GRAY80, COLOR_DEFAULT_FOREGROUND = 0x0040,
              COLOR_DEFAULT_BACKGROUND = 0x0041, COLOR_TOOLTIP = 0x0051,
              COLOR_AUTO = 0x7FFF} Color;

typedef enum {NUMFORMAT_GENERAL, NUMFORMAT_NUMBER, NUMFORMAT_NUMBER_D2,
              NUMFORMAT_NUMBER_SEP, NUMFORMAT_NUMBER_SEP_D2,
              NUMFORMAT_CURRENCY_NEGBRA, NUMFORMAT_CURRENCY_NEGBRARED,
              NUMFORMAT_CURRENCY_D2_NEGBRA, NUMFORMAT_CURRENCY_D2_NEGBRARED,
              NUMFORMAT_PERCENT, NUMFORMAT_PERCENT_D2, NUMFORMAT_SCIENTIFIC_D2,
              NUMFORMAT_FRACTION_ONEDIG, NUMFORMAT_FRACTION_TWODIG,
              NUMFORMAT_DATE, NUMFORMAT_CUSTOM_D_MON_YY,
              NUMFORMAT_CUSTOM_D_MON, NUMFORMAT_CUSTOM_MON_YY,
              NUMFORMAT_CUSTOM_HMM_AM, NUMFORMAT_CUSTOM_HMMSS_AM,
              NUMFORMAT_CUSTOM_HMM, NUMFORMAT_CUSTOM_HMMSS,
              NUMFORMAT_CUSTOM_MDYYYY_HMM, NUMFORMAT_NUMBER_SEP_NEGBRA = 37,
              NUMFORMAT_NUMBER_SEP_NEGBRARED, NUMFORMAT_NUMBER_D2_SEP_NEGBRA,
              NUMFORMAT_NUMBER_D2_SEP_NEGBRARED, NUMFORMAT_ACCOUNT,
              NUMFORMAT_ACCOUNTCUR, NUMFORMAT_ACCOUNT_D2,
              NUMFORMAT_ACCOUNT_D2_CUR, NUMFORMAT_CUSTOM_MMSS,
              NUMFORMAT_CUSTOM_H0MMSS, NUMFORMAT_CUSTOM_MMSS0,
              NUMFORMAT_CUSTOM_000P0E_PLUS0, NUMFORMAT_TEXT} NumFormat;

typedef enum {ALIGNH_GENERAL, ALIGNH_LEFT, ALIGNH_CENTER, ALIGNH_RIGHT,
              ALIGNH_FILL, ALIGNH_JUSTIFY, ALIGNH_MERGE, ALIGNH_DISTRIBUTED}
              AlignH;

typedef enum {ALIGNV_TOP, ALIGNV_CENTER, ALIGNV_BOTTOM, ALIGNV_JUSTIFY,
              ALIGNV_DISTRIBUTED} AlignV;

typedef enum {BORDERSTYLE_NONE, BORDERSTYLE_THIN, BORDERSTYLE_MEDIUM,
              BORDERSTYLE_DASHED, BORDERSTYLE_DOTTED, BORDERSTYLE_THICK,
              BORDERSTYLE_DOUBLE, BORDERSTYLE_HAIR, BORDERSTYLE_MEDIUMDASHED,
              BORDERSTYLE_DASHDOT, BORDERSTYLE_MEDIUMDASHDOT,
              BORDERSTYLE_DASHDOTDOT, BORDERSTYLE_MEDIUMDASHDOTDOT,
              BORDERSTYLE_SLANTDASHDOT} BorderStyle;

typedef enum {BORDERDIAGONAL_NONE, BORDERDIAGONAL_DOWN, BORDERDIAGONAL_UP,
              BORDERDIAGONAL_BOTH} BorderDiagonal;

typedef enum {FILLPATTERN_NONE, FILLPATTERN_SOLID, FILLPATTERN_GRAY50,
              FILLPATTERN_GRAY75, FILLPATTERN_GRAY25, FILLPATTERN_HORSTRIPE,
              FILLPATTERN_VERSTRIPE, FILLPATTERN_REVDIAGSTRIPE,
              FILLPATTERN_DIAGSTRIPE, FILLPATTERN_DIAGCROSSHATCH,
              FILLPATTERN_THICKDIAGCROSSHATCH, FILLPATTERN_THINHORSTRIPE,
              FILLPATTERN_THINVERSTRIPE, FILLPATTERN_THINREVDIAGSTRIPE,
              FILLPATTERN_THINDIAGSTRIPE, FILLPATTERN_THINHORCROSSHATCH,
              FILLPATTERN_THINDIAGCROSSHATCH, FILLPATTERN_GRAY12P5,
              FILLPATTERN_GRAY6P25} FillPattern;

typedef enum {SCRIPT_NORMAL, SCRIPT_SUPER, SCRIPT_SUB} Script;

typedef enum {UNDERLINE_NONE, UNDERLINE_SINGLE, UNDERLINE_DOUBLE,
              UNDERLINE_SINGLEACC = 0x21, UNDERLINE_DOUBLEACC = 0x22} Underline;

typedef enum {PAPER_DEFAULT, PAPER_LETTER, PAPER_LETTERSMALL, PAPER_TABLOID,
              PAPER_LEDGER, PAPER_LEGAL, PAPER_STATEMENT, PAPER_EXECUTIVE,
              PAPER_A3, PAPER_A4, PAPER_A4SMALL, PAPER_A5, PAPER_B4, PAPER_B5,
              PAPER_FOLIO, PAPER_QUATRO, PAPER_10x14, PAPER_10x17, PAPER_NOTE,
              PAPER_ENVELOPE_9, PAPER_ENVELOPE_10, PAPER_ENVELOPE_11,
              PAPER_ENVELOPE_12, PAPER_ENVELOPE_14, PAPER_C_SIZE, PAPER_D_SIZE,
              PAPER_E_SIZE, PAPER_ENVELOPE_DL, PAPER_ENVELOPE_C5,
              PAPER_ENVELOPE_C3, PAPER_ENVELOPE_C4, PAPER_ENVELOPE_C6,
              PAPER_ENVELOPE_C65, PAPER_ENVELOPE_B4, PAPER_ENVELOPE_B5,
              PAPER_ENVELOPE_B6, PAPER_ENVELOPE, PAPER_ENVELOPE_MONARCH,
              PAPER_US_ENVELOPE, PAPER_FANFOLD, PAPER_GERMAN_STD_FANFOLD,
              PAPER_GERMAN_LEGAL_FANFOLD, PAPER_B4_ISO,
              PAPER_JAPANESE_POSTCARD, PAPER_9x11, PAPER_10x11, PAPER_15x11,
              PAPER_ENVELOPE_INVITE, PAPER_US_LETTER_EXTRA = 50,
              PAPER_US_LEGAL_EXTRA, PAPER_US_TABLOID_EXTRA, PAPER_A4_EXTRA,
              PAPER_LETTER_TRANSVERSE, PAPER_A4_TRANSVERSE,
              PAPER_LETTER_EXTRA_TRANSVERSE, PAPER_SUPERA, PAPER_SUPERB,
              PAPER_US_LETTER_PLUS, PAPER_A4_PLUS, PAPER_A5_TRANSVERSE,
              PAPER_B5_TRANSVERSE, PAPER_A3_EXTRA, PAPER_A5_EXTRA,
              PAPER_B5_EXTRA, PAPER_A2, PAPER_A3_TRANSVERSE,
              PAPER_A3_EXTRA_TRANSVERSE, PAPER_JAPANESE_DOUBLE_POSTCARD,
              PAPER_A6, PAPER_JAPANESE_ENVELOPE_KAKU2,
              PAPER_JAPANESE_ENVELOPE_KAKU3, PAPER_JAPANESE_ENVELOPE_CHOU3,
              PAPER_JAPANESE_ENVELOPE_CHOU4, PAPER_LETTER_ROTATED,
              PAPER_A3_ROTATED, PAPER_A4_ROTATED, PAPER_A5_ROTATED,
              PAPER_B4_ROTATED, PAPER_B5_ROTATED,
              PAPER_JAPANESE_POSTCARD_ROTATED,
              PAPER_DOUBLE_JAPANESE_POSTCARD_ROTATED, PAPER_A6_ROTATED,
              PAPER_JAPANESE_ENVELOPE_KAKU2_ROTATED,
              PAPER_JAPANESE_ENVELOPE_KAKU3_ROTATED,
              PAPER_JAPANESE_ENVELOPE_CHOU3_ROTATED,
              PAPER_JAPANESE_ENVELOPE_CHOU4_ROTATED, PAPER_B6,
              PAPER_B6_ROTATED, PAPER_12x11, PAPER_JAPANESE_ENVELOPE_YOU4,
              PAPER_JAPANESE_ENVELOPE_YOU4_ROTATED, PAPER_PRC16K, PAPER_PRC32K,
              PAPER_PRC32K_BIG, PAPER_PRC_ENVELOPE1, PAPER_PRC_ENVELOPE2,
              PAPER_PRC_ENVELOPE3, PAPER_PRC_ENVELOPE4, PAPER_PRC_ENVELOPE5,
              PAPER_PRC_ENVELOPE6, PAPER_PRC_ENVELOPE7, PAPER_PRC_ENVELOPE8,
              PAPER_PRC_ENVELOPE9, PAPER_PRC_ENVELOPE10, PAPER_PRC16K_ROTATED,
              PAPER_PRC32K_ROTATED, PAPER_PRC32KBIG_ROTATED,
              PAPER_PRC_ENVELOPE1_ROTATED, PAPER_PRC_ENVELOPE2_ROTATED,
              PAPER_PRC_ENVELOPE3_ROTATED, PAPER_PRC_ENVELOPE4_ROTATED,
              PAPER_PRC_ENVELOPE5_ROTATED, PAPER_PRC_ENVELOPE6_ROTATED,
              PAPER_PRC_ENVELOPE7_ROTATED, PAPER_PRC_ENVELOPE8_ROTATED,
              PAPER_PRC_ENVELOPE9_ROTATED, PAPER_PRC_ENVELOPE10_ROTATED} Paper;

typedef enum {SHEETTYPE_SHEET, SHEETTYPE_CHART, SHEETTYPE_UNKNOWN} SheetType;

typedef enum {CELLTYPE_EMPTY, CELLTYPE_NUMBER, CELLTYPE_STRING,
              CELLTYPE_BOOLEAN, CELLTYPE_BLANK, CELLTYPE_ERROR} CellType;

typedef enum {ERRORTYPE_NULL = 0x00, ERRORTYPE_DIV_0 = 0x07,
              ERRORTYPE_VALUE = 0x0F, ERRORTYPE_REF = 0x17,
              ERRORTYPE_NAME = 0x1D, ERRORTYPE_NUM = 0x24,
              ERRORTYPE_NA = 0x2A} ErrorType;

typedef enum {SHEETSTATE_VISIBLE, SHEETSTATE_HIDDEN, SHEETSTATE_VERYHIDDEN}
              SheetState;

typedef struct tagBookHandle   * BookHandle;
typedef struct tagSheetHandle  * SheetHandle;
typedef struct tagFormatHandle * FormatHandle;
typedef struct tagFontHandle   * FontHandle;

BookHandle   __cdecl xlCreateBookCA();
BookHandle   __cdecl xlCreateXMLBookCA();

int          __cdecl xlBookLoadA(BookHandle handle, const char* filename);
int          __cdecl xlBookSaveA(BookHandle handle, const char* filename);
int          __cdecl xlBookLoadRawA(BookHandle handle, const char* data,
                                    unsigned size);
int          __cdecl xlBookSaveRawA(BookHandle handle, const char** data,
                                    unsigned* size);
SheetHandle  __cdecl xlBookAddSheetA(BookHandle handle, const char* name,
                                     SheetHandle initSheet);
SheetHandle  __cdecl xlBookInsertSheetA(BookHandle handle, int index,
                                        const char* name,
                                        SheetHandle initSheet);
SheetHandle  __cdecl xlBookGetSheetA(BookHandle handle, int index);
int          __cdecl xlBookSheetTypeA(BookHandle handle, int index);
int          __cdecl xlBookDelSheetA(BookHandle handle, int index);
int          __cdecl xlBookSheetCountA(BookHandle handle);
FormatHandle __cdecl xlBookAddFormatA(BookHandle handle,
                                      FormatHandle initFormat);
FontHandle   __cdecl xlBookAddFontA(BookHandle handle, FontHandle initFont);
int          __cdecl xlBookAddCustomNumFormatA(BookHandle handle,
                                               const char* customNumFormat);
const char*  __cdecl xlBookCustomNumFormatA(BookHandle handle, int fmt);
FormatHandle __cdecl xlBookFormatA(BookHandle handle, int index);
int          __cdecl xlBookFormatSizeA(BookHandle handle);
FontHandle   __cdecl xlBookFontA(BookHandle handle, int index);
int          __cdecl xlBookFontSizeA(BookHandle handle);
double       __cdecl xlBookDatePackA(BookHandle handle, int year, int month,
                                     int day, int hour, int min, int sec,
                                     int msec);
int          __cdecl xlBookDateUnpackA(BookHandle handle, double value,
                                       int* year, int* month, int* day,
                                       int* hour, int* min, int* sec,
                                       int* msec);
int          __cdecl xlBookColorPackA(BookHandle handle, int red, int green,
                                      int blue);
void         __cdecl xlBookColorUnpackA(BookHandle handle, int color, int* red,
                                        int* green, int* blue);
int          __cdecl xlBookActiveSheetA(BookHandle handle);
void         __cdecl xlBookSetActiveSheetA(BookHandle handle, int index);
int          __cdecl xlBookPictureSizeA(BookHandle handle);
int          __cdecl xlBookGetPictureA(BookHandle handle, int index,
                                       const char** data, unsigned* size);
int          __cdecl xlBookAddPictureA(BookHandle handle, const char* filename);
int          __cdecl xlBookAddPicture2A(BookHandle handle, const char* data,
                                        unsigned size);
const char*  __cdecl xlBookDefaultFontA(BookHandle handle, int* fontSize);
void         __cdecl xlBookSetDefaultFontA(BookHandle handle,
                                           const char* fontName, int fontSize);
int          __cdecl xlBookRefR1C1A(BookHandle handle);
void         __cdecl xlBookSetRefR1C1A(BookHandle handle, int refR1C1);
void         __cdecl xlBookSetKeyA(BookHandle handle, const char* name,
                                   const char* key);
int          __cdecl xlBookRgbModeA(BookHandle handle);
void         __cdecl xlBookSetRgbModeA(BookHandle handle, int rgbMode);
int          __cdecl xlBookVersionA(BookHandle handle);
int          __cdecl xlBookBiffVersionA(BookHandle handle);
int          __cdecl xlBookIsDate1904A(BookHandle handle);
void         __cdecl xlBookSetDate1904A(BookHandle handle, int date1904);
int          __cdecl xlBookSetLocaleA(BookHandle handle, const char* locale);
const char*  __cdecl xlBookErrorMessageA(BookHandle handle);
void         __cdecl xlBookReleaseA(BookHandle handle);

int          __cdecl xlSheetCellTypeA(SheetHandle handle, int row, int col);
int          __cdecl xlSheetIsFormulaA(SheetHandle handle, int row, int col);
FormatHandle __cdecl xlSheetCellFormatA(SheetHandle handle, int row, int col);
void         __cdecl xlSheetSetCellFormatA(SheetHandle handle, int row,
                                           int col, FormatHandle format);
const char*  __cdecl xlSheetReadStrA(SheetHandle handle, int row, int col,
                                     FormatHandle* format);
int          __cdecl xlSheetWriteStrA(SheetHandle handle, int row, int col,
                                      const char* value, FormatHandle format);
double       __cdecl xlSheetReadNumA(SheetHandle handle, int row, int col,
                                     FormatHandle* format);
int          __cdecl xlSheetWriteNumA(SheetHandle handle, int row, int col,
                                      double value, FormatHandle format);
int          __cdecl xlSheetReadBoolA(SheetHandle handle, int row, int col,
                                      FormatHandle* format);
int          __cdecl xlSheetWriteBoolA(SheetHandle handle, int row, int col,
                                       int value, FormatHandle format);
int          __cdecl xlSheetReadBlankA(SheetHandle handle, int row, int col,
                                       FormatHandle* format);
int          __cdecl xlSheetWriteBlankA(SheetHandle handle, int row, int col,
                                        FormatHandle format);
const char*  __cdecl xlSheetReadFormulaA(SheetHandle handle, int row, int col,
                                         FormatHandle* format);
int          __cdecl xlSheetWriteFormulaA(SheetHandle handle, int row, int col,
                                          const char* value,
                                          FormatHandle format);
const char*  __cdecl xlSheetReadCommentA(SheetHandle handle, int row, int col);
void         __cdecl xlSheetWriteCommentA(SheetHandle handle, int row, int col,
                                          const char* value, const char* author,
                                          int width, int height);
int          __cdecl xlSheetIsDateA(SheetHandle handle, int row, int col);
int          __cdecl xlSheetReadErrorA(SheetHandle handle, int row, int col);
double       __cdecl xlSheetColWidthA(SheetHandle handle, int col);
double       __cdecl xlSheetRowHeightA(SheetHandle handle, int row);
int          __cdecl xlSheetSetColA(SheetHandle handle, int colFirst, int
                                    colLast, double width, FormatHandle format,
                                    int hidden);
int          __cdecl xlSheetSetRowA(SheetHandle handle, int row, double height,
                                    FormatHandle format, int hidden);
int          __cdecl xlSheetRowHiddenA(SheetHandle handle, int row);
int          __cdecl xlSheetSetRowHiddenA(SheetHandle handle, int row,
                                          int hidden);
int          __cdecl xlSheetColHiddenA(SheetHandle handle, int col);
int          __cdecl xlSheetSetColHiddenA(SheetHandle handle, int col,
                                          int hidden);
int          __cdecl xlSheetGetMergeA(SheetHandle handle, int row, int col,
                                      int* rowFirst, int* rowLast,
                                      int* colFirst, int* colLast);
int          __cdecl xlSheetSetMergeA(SheetHandle handle, int rowFirst,
                                      int rowLast, int colFirst, int colLast);
int          __cdecl xlSheetDelMergeA(SheetHandle handle, int row, int col);
int          __cdecl xlSheetPictureSizeA(SheetHandle handle);
int          __cdecl xlSheetGetPictureA(SheetHandle handle, int index,
                                        int* rowTop, int* colLeft,
                                        int* rowBottom, int* colRight,
                                        int* width, int* height, int* offset_x,
                                        int* offset_y);
void         __cdecl xlSheetSetPictureA(SheetHandle handle, int row, int col,
                                        int pictureId, double scale,
                                        int offset_x, int offset_y);
void         __cdecl xlSheetSetPicture2A(SheetHandle handle, int row, int col,
                                         int pictureId, int width, int height,
                                         int offset_x, int offset_y);
int          __cdecl xlSheetGetHorPageBreakA(SheetHandle handle, int index);
int          __cdecl xlSheetGetHorPageBreakSizeA(SheetHandle handle);
int          __cdecl xlSheetGetVerPageBreakA(SheetHandle handle, int index);
int          __cdecl xlSheetGetVerPageBreakSizeA(SheetHandle handle);
int          __cdecl xlSheetSetHorPageBreakA(SheetHandle handle, int row,
                                             int pageBreak);
int          __cdecl xlSheetSetVerPageBreakA(SheetHandle handle, int col,
                                             int pageBreak);
void         __cdecl xlSheetSplitA(SheetHandle handle, int row, int col);
int          __cdecl xlSheetGroupRowsA(SheetHandle handle, int rowFirst,
                                       int rowLast, int collapsed);
int          __cdecl xlSheetGroupColsA(SheetHandle handle, int colFirst,
                                       int colLast, int collapsed);
int          __cdecl xlSheetGroupSummaryBelowA(SheetHandle handle);
void         __cdecl xlSheetSetGroupSummaryBelowA(SheetHandle handle,
                                                  int below);
int          __cdecl xlSheetGroupSummaryRightA(SheetHandle handle);
void         __cdecl xlSheetSetGroupSummaryRightA(SheetHandle handle,
                                                  int right);
void         __cdecl xlSheetClearA(SheetHandle handle, int rowFirst,
                                   int rowLast, int colFirst, int colLast);
int          __cdecl xlSheetInsertColA(SheetHandle handle, int colFirst,
                                       int colLast);
int          __cdecl xlSheetInsertRowA(SheetHandle handle, int rowFirst,
                                       int rowLast);
int          __cdecl xlSheetRemoveColA(SheetHandle handle, int colFirst,
                                       int colLast);
int          __cdecl xlSheetRemoveRowA(SheetHandle handle, int rowFirst,
                                       int rowLast);
int          __cdecl xlSheetCopyCellA(SheetHandle handle, int rowSrc,
                                      int colSrc, int rowDst, int colDst);
int          __cdecl xlSheetFirstRowA(SheetHandle handle);
int          __cdecl xlSheetLastRowA(SheetHandle handle);
int          __cdecl xlSheetFirstColA(SheetHandle handle);
int          __cdecl xlSheetLastColA(SheetHandle handle);
int          __cdecl xlSheetDisplayGridlinesA(SheetHandle handle);
void         __cdecl xlSheetSetDisplayGridlinesA(SheetHandle handle, int show);
int          __cdecl xlSheetPrintGridlinesA(SheetHandle handle);
void         __cdecl xlSheetSetPrintGridlinesA(SheetHandle handle, int print);
int          __cdecl xlSheetZoomA(SheetHandle handle);
void         __cdecl xlSheetSetZoomA(SheetHandle handle, int zoom);
int          __cdecl xlSheetPrintZoomA(SheetHandle handle);
void         __cdecl xlSheetSetPrintZoomA(SheetHandle handle, int zoom);
int          __cdecl xlSheetGetPrintFitA(SheetHandle handle, int* wPages,
                                         int* hPages);
void         __cdecl xlSheetSetPrintFitA(SheetHandle handle, int wPages,
                                         int hPages);
int          __cdecl xlSheetLandscapeA(SheetHandle handle);
void         __cdecl xlSheetSetLandscapeA(SheetHandle handle, int landscape);
int          __cdecl xlSheetPaperA(SheetHandle handle);
void         __cdecl xlSheetSetPaperA(SheetHandle handle, int paper);
const char*  __cdecl xlSheetHeaderA(SheetHandle handle);
int          __cdecl xlSheetSetHeaderA(SheetHandle handle, const char* header,
                                       double margin);
double       __cdecl xlSheetHeaderMarginA(SheetHandle handle);
const char*  __cdecl xlSheetFooterA(SheetHandle handle);
int          __cdecl xlSheetSetFooterA(SheetHandle handle, const char* footer,
                                       double margin);
double       __cdecl xlSheetFooterMarginA(SheetHandle handle);
int          __cdecl xlSheetHCenterA(SheetHandle handle);
void         __cdecl xlSheetSetHCenterA(SheetHandle handle, int hCenter);
int          __cdecl xlSheetVCenterA(SheetHandle handle);
void         __cdecl xlSheetSetVCenterA(SheetHandle handle, int vCenter);
double       __cdecl xlSheetMarginLeftA(SheetHandle handle);
void         __cdecl xlSheetSetMarginLeftA(SheetHandle handle, double margin);
double       __cdecl xlSheetMarginRightA(SheetHandle handle);
void         __cdecl xlSheetSetMarginRightA(SheetHandle handle, double margin);
double       __cdecl xlSheetMarginTopA(SheetHandle handle);
void         __cdecl xlSheetSetMarginTopA(SheetHandle handle, double margin);
double       __cdecl xlSheetMarginBottomA(SheetHandle handle);
void         __cdecl xlSheetSetMarginBottomA(SheetHandle handle, double margin);
int          __cdecl xlSheetPrintRowColA(SheetHandle handle);
void         __cdecl xlSheetSetPrintRowColA(SheetHandle handle, int print);
void         __cdecl xlSheetSetPrintRepeatRowsA(SheetHandle handle,
                                                int rowFirst, int rowLast);
void         __cdecl xlSheetSetPrintRepeatColsA(SheetHandle handle,
                                                int colFirst, int colLast);
void         __cdecl xlSheetSetPrintAreaA(SheetHandle handle, int rowFirst,
                                          int rowLast, int colFirst,
                                          int colLast);
void         __cdecl xlSheetClearPrintRepeatsA(SheetHandle handle);
void         __cdecl xlSheetClearPrintAreaA(SheetHandle handle);
int          __cdecl xlSheetGetNamedRangeA(SheetHandle handle,
                                           const char* name, int* rowFirst,
                                           int* rowLast, int* colFirst,
                                           int* colLast);
int          __cdecl xlSheetSetNamedRangeA(SheetHandle handle,
                                           const char* name, int rowFirst,
                                           int rowLast, int colFirst,
                                           int colLast);
int          __cdecl xlSheetDelNamedRangeA(SheetHandle handle,
                                           const char* name);
int          __cdecl xlSheetNamedRangeSizeA(SheetHandle handle);
const char*  __cdecl xlSheetNamedRangeA(SheetHandle handle, int index,
                                        int* rowFirst, int* rowLast,
                                        int* colFirst, int* colLast);
const char*  __cdecl xlSheetNameA(SheetHandle handle);
void         __cdecl xlSheetSetNameA(SheetHandle handle, const char* name);
int          __cdecl xlSheetProtectA(SheetHandle handle);
void         __cdecl xlSheetSetProtectA(SheetHandle handle, int protect);
int          __cdecl xlSheetHiddenA(SheetHandle handle);
int          __cdecl xlSheetSetHiddenA(SheetHandle handle, int hidden);
void         __cdecl xlSheetGetTopLeftViewA(SheetHandle handle, int* row,
                                            int* col);
void         __cdecl xlSheetSetTopLeftViewA(SheetHandle handle, int row,
                                            int col);
void         __cdecl xlSheetAddrToRowColA(SheetHandle handle, const char* addr,
                                          int* row, int* col, int* rowRelative,
                                          int* colRelative);
const char*  __cdecl xlSheetRowColToAddrA(SheetHandle handle, int row, int col,
                                          int rowRelative, int colRelative);

FontHandle   __cdecl xlFormatFontA(FormatHandle handle);
int          __cdecl xlFormatSetFontA(FormatHandle handle,
                                      FontHandle fontHandle);
int          __cdecl xlFormatNumFormatA(FormatHandle handle);
void         __cdecl xlFormatSetNumFormatA(FormatHandle handle, int numFormat);
int          __cdecl xlFormatAlignHA(FormatHandle handle);
void         __cdecl xlFormatSetAlignHA(FormatHandle handle, int align);
int          __cdecl xlFormatAlignVA(FormatHandle handle);
void         __cdecl xlFormatSetAlignVA(FormatHandle handle, int align);
int          __cdecl xlFormatWrapA(FormatHandle handle);
void         __cdecl xlFormatSetWrapA(FormatHandle handle, int wrap);
int          __cdecl xlFormatRotationA(FormatHandle handle);
int          __cdecl xlFormatSetRotationA(FormatHandle handle, int rotation);
int          __cdecl xlFormatIndentA(FormatHandle handle);
void         __cdecl xlFormatSetIndentA(FormatHandle handle, int indent);
int          __cdecl xlFormatShrinkToFitA(FormatHandle handle);
void         __cdecl xlFormatSetShrinkToFitA(FormatHandle handle,
                                             int shrinkToFit);
void         __cdecl xlFormatSetBorderA(FormatHandle handle, int style);
void         __cdecl xlFormatSetBorderColorA(FormatHandle handle, int color);
int          __cdecl xlFormatBorderLeftA(FormatHandle handle);
void         __cdecl xlFormatSetBorderLeftA(FormatHandle handle, int style);
int          __cdecl xlFormatBorderRightA(FormatHandle handle);
void         __cdecl xlFormatSetBorderRightA(FormatHandle handle, int style);
int          __cdecl xlFormatBorderTopA(FormatHandle handle);
void         __cdecl xlFormatSetBorderTopA(FormatHandle handle, int style);
int          __cdecl xlFormatBorderBottomA(FormatHandle handle);
void         __cdecl xlFormatSetBorderBottomA(FormatHandle handle, int style);
int          __cdecl xlFormatBorderLeftColorA(FormatHandle handle);
void         __cdecl xlFormatSetBorderLeftColorA(FormatHandle handle,
                                                 int color);
int          __cdecl xlFormatBorderRightColorA(FormatHandle handle);
void         __cdecl xlFormatSetBorderRightColorA(FormatHandle handle,
                                                  int color);
int          __cdecl xlFormatBorderTopColorA(FormatHandle handle);
void         __cdecl xlFormatSetBorderTopColorA(FormatHandle handle,
                                                int color);
int          __cdecl xlFormatBorderBottomColorA(FormatHandle handle);
void         __cdecl xlFormatSetBorderBottomColorA(FormatHandle handle,
                                                   int color);
int          __cdecl xlFormatBorderDiagonalA(FormatHandle handle);
void         __cdecl xlFormatSetBorderDiagonalA(FormatHandle handle,
                                                int border);
int          __cdecl xlFormatBorderDiagonalStyleA(FormatHandle handle);
void         __cdecl xlFormatSetBorderDiagonalStyleA(FormatHandle handle,
                                                     int style);
int          __cdecl xlFormatBorderDiagonalColorA(FormatHandle handle);
void         __cdecl xlFormatSetBorderDiagonalColorA(FormatHandle handle,
                                                     int color);
int          __cdecl xlFormatFillPatternA(FormatHandle handle);
void         __cdecl xlFormatSetFillPatternA(FormatHandle handle, int pattern);
int          __cdecl xlFormatPatternForegroundColorA(FormatHandle handle);
void         __cdecl xlFormatSetPatternForegroundColorA(FormatHandle handle,
                                                        int color);
int          __cdecl xlFormatPatternBackgroundColorA(FormatHandle handle);
void         __cdecl xlFormatSetPatternBackgroundColorA(FormatHandle handle,
                                                        int color);
int          __cdecl xlFormatLockedA(FormatHandle handle);
void         __cdecl xlFormatSetLockedA(FormatHandle handle, int locked);
int          __cdecl xlFormatHiddenA(FormatHandle handle);
void         __cdecl xlFormatSetHiddenA(FormatHandle handle, int hidden);

int          __cdecl xlFontSizeA(FontHandle handle);
void         __cdecl xlFontSetSizeA(FontHandle handle, int size);
int          __cdecl xlFontItalicA(FontHandle handle);
void         __cdecl xlFontSetItalicA(FontHandle handle, int italic);
int          __cdecl xlFontStrikeOutA(FontHandle handle);
void         __cdecl xlFontSetStrikeOutA(FontHandle handle, int strikeOut);
int          __cdecl xlFontColorA(FontHandle handle);
void         __cdecl xlFontSetColorA(FontHandle handle, int color);
int          __cdecl xlFontBoldA(FontHandle handle);
void         __cdecl xlFontSetBoldA(FontHandle handle, int bold);
int          __cdecl xlFontScriptA(FontHandle handle);
void         __cdecl xlFontSetScriptA(FontHandle handle, int script);
int          __cdecl xlFontUnderlineA(FontHandle handle);
void         __cdecl xlFontSetUnderlineA(FontHandle handle, int underline);
const char*  __cdecl xlFontNameA(FontHandle handle);
void         __cdecl xlFontSetNameA(FontHandle handle, const char* name);
]]

local year  = ffi_new("int[1]", 0)
local month = ffi_new("int[1]", 0)
local day   = ffi_new("int[1]", 0)
local hour  = ffi_new("int[1]", 0)
local min   = ffi_new("int[1]", 0)
local sec   = ffi_new("int[1]", 0)
local msec  = ffi_new("int[1]", 0)

local book = { sheets = {} }
local sheet = {}

function book:new(opts)
    opts = opts or {}
    local o = { sheets = {} }
    setmetatable(o, self)
    setmetatable(o.sheets, self.sheets)
    self.__index = self
    self.sheets.__index = self.sheets
    if opts.format == "xls" then
        o.___ = libxl.xlCreateBookCA()
    else
        o.___ = libxl.xlCreateXMLBookCA()
    end
    o.sheets.book = o
    return o
end

function book:load(filename)
    libxl.xlBookLoadA(self.___, filename)
    local count = libxl.xlBookSheetCountA(self.___) - 1
    for i=0,count do
        self.sheets[#self.sheets + 1] = sheet:new(self, libxl.xlBookGetSheetA(self.___, i))
    end
    return self
end

function book:save(filename, release)
    libxl.xlBookSaveA(self.___, filename)
    if release then self:release() end
    return self
end

function book:release()
    libxl.xlBookReleaseA(self.___)
    return self
end

function book:date_unpack(value)
    libxl.xlBookDateUnpackA(self.___,
        value, year[0], month[0], day[0],
        hour[0], min[0], sec[0], msec[0])
    return  year[0], month[0], day[0], hour[0], min[0], sec[0], msec[0]
end

function book.sheets:add(name, initSheet)
    if type(initSheet) ~= "table" then initSheet = {} end
    self[#self + 1] = sheet:new(self.book, libxl.xlBookAddSheetA(
        self.book.___, name, initSheet.___ or nil))
    return self[#self]
end

function book.sheets:insert(index, name, initSheet)
    if type(initSheet) ~= "table" then initSheet = {} end
    table.insert(self, sheet:new(self.book, libxl.xlBookInsertSheetA(
        self.book.___, index - 1, name, initSheet.___ or nil)), index)
    return self[index]
end

function sheet:new(book, ___)
    local o = { book = book, ___ = ___ }
    setmetatable(o, self)
    self.__index = self
    return o
end

function sheet:read(row, col, format)
    if self:is_formula(row, col) then
        return ffi_string(libxl.xlSheetReadFormulaA(self.___, row, col, format))
    elseif self:is_date(row, col) then
        return self.book:unpack_date(
            libxl.xlSheetReadNumA(self.___, row, col, format)
        )
    else
        local  t = libxl.xlSheetCellTypeA(self.___, row, col)
        if     t == C.CELLTYPE_EMPTY   then
            return ""
        elseif t == C.CELLTYPE_NUMBER  then
            return libxl.xlSheetReadNumA(self.___, row, col, format)
        elseif t == C.CELLTYPE_STRING  then
            return ffi_string(libxl.xlSheetReadStrA(self.___, row, col, format))
        elseif t == C.CELLTYPE_BOOLEAN then
            return libxl.xlSheetReadBoolA(self.___, row, col, format) ~= 0
        elseif t == C.CELLTYPE_BLANK   then
            libxl.xlSheetReadBlankA(self.___, row, col, format)
            return nil
        elseif t == C.CELLTYPE_ERROR   then
            return libxl.xlSheetReadErrorA(self.___, row, col)
        else
            return nil
        end
    end
end

function sheet:write(row, col, value, format)

    if value == nil then
        libxl.xlSheetWriteBlankA(self.___, row, col, format)
    else
        local  t = type(value)
        if     t == "string" then
            libxl.xlSheetWriteStrA(self.___, row, col, value, format)
        elseif t == "number" then
            libxl.xlSheetWriteNumA(self.___, row, col, value, format)
        elseif t == "boolean" then
            libxl.xlSheetWriteBoolA(self.___, row, col, value, format)
        end
    end
    return self
end

function sheet:is_formula(row, col)
    return libxl.xlSheetIsFormulaA(self.___, row, col) ~= 0
end

function sheet:is_date(row, col)
    return libxl.xlSheetIsDateA(self.___, row, col) ~= 0
end

function sheet:cell_type(row, col)
    local  t = libxl.xlSheetCellTypeA(self.___, row, col)
    if     t == C.CELLTYPE_EMPTY   then
    elseif t == C.CELLTYPE_NUMBER  then
    elseif t == C.CELLTYPE_STRING  then
    elseif t == C.CELLTYPE_BOOLEAN then
    elseif t == C.CELLTYPE_BLANK   then
    elseif t == C.CELLTYPE_ERROR   then
    else
    end
end

function sheet:__newindex(n, v)
    if (n == "name") then
        libxl.xlSheetSetNameA(self.___, v)
    else
        rawset(self, n, v)
    end
end

return book

