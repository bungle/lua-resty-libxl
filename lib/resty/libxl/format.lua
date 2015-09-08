require "resty.libxl.types.handle"
require "resty.libxl.types.format"
local lib          = require "resty.libxl.library"
local setmetatable = setmetatable
local rawget       = rawget
local rawset       = rawset

local format = {}

function format.new(opts)
    return setmetatable(opts, format)
end

function format:__index(n)
    if n == "font" then
        return font.new{ context = lib.xlFormatFontA(self.context) }
    elseif n == "num-format" then
        return lib.xlFormatNumFormatA(self.context)
    elseif n == "align-h" then
        return lib.xlFormatAlignHA(self.context)
    elseif n == "align-v" then
        return lib.xlFormatAlignVA(self.context)
    elseif n == "wrap" then
        return lib.xlFormatWrapA(self.context) == 1
    elseif n == "rotation" then
        return lib.xlFormatRotationA(self.context)
    elseif n == "indent" then
        return lib.xlFormatIndentA(self.context)
    elseif n == "shrink-to-fit" then
        return lib.xlFormatShrinkToFitA(self.context)
    elseif n == "border-left" then
        return lib.xlFormatBorderLeftA(self.context)
    elseif n == "border-right" then
        return lib.xlFormatBorderRightA(self.context)
    elseif n == "border-top" then
        return lib.xlFormatBorderTopA(self.context)
    elseif n == "border-bottom" then
        return lib.xlFormatBorderBottomA(self.context)
    elseif n == "border-left-color" then
        return lib.xlFormatBorderLeftColorA(self.context)
    elseif n == "border-right-color" then
        return lib.xlFormatBorderRightColorA(self.context)
    elseif n == "border-top-color" then
        return lib.xlFormatBorderTopColorA(self.context)
    elseif n == "border-bottom-color" then
        return lib.xlFormatBorderBottomColorA(self.context)
    elseif n == "border-diagonal" then
        return lib.xlFormatBorderDiagonalA(self.context)
    elseif n == "border-diagonal-style" then
        return lib.xlFormatBorderDiagonalStyleA(self.context)
    elseif n == "border-diagonal-color" then
        return lib.xlFormatBorderDiagonalColorA(self.context)
    elseif n == "fill-pattern" then
        return lib.xlFormatFillPatternA(self.context)
    elseif n == "pattern-foreground-color" then
        return lib.xlFormatPatternForegroundColorA(self.context)
    elseif n == "pattern-background-color" then
        return lib.xlFormatPatternBackgroundColorA(self.context)
    elseif n == "locked" then
        return lib.xlFormatLockedA(self.context) == 1
    elseif n == "hidden" then
        return lib.xlFormatHiddenA(self.context) == 1
    else
        rawget(format, n)
    end
end

function format:__newindex(n, v)
    if n == "font" then
        lib.xlFormatSetNumFormatA(self.context, v.context or v)
    elseif n == "num-format" then
        lib.xlFormatSetNumFormatA(self.context, v)
    elseif n == "align-h" then
        lib.xlFormatSetAlignHA(self.context, v)
    elseif n == "align-v" then
        lib.xlFormatSetAlignVA(self.context, v)
    elseif n == "wrap" then
        lib.xlFormatSetWrapA(self.context, v)
    elseif n == "rotation" then
        lib.xlFormatSetRotationA(self.context, v)
    elseif n == "indent" then
        lib.xlFormatSetIndentA(self.context, v)
    elseif n == "shrink-to-fit" then
        lib.xlFormatSetShrinkToFitA(self.context, v)
    elseif n == "border" then
        lib.xlFormatSetBorderA(self.context, v)
    elseif n == "border-color" then
        lib.xlFormatSetBorderColorA(self.context, v)
    elseif n == "border-left" then
        lib.xlFormatSetBorderLeftA(self.context, v)
    elseif n == "border-right" then
        lib.xlFormatSetBorderRightA(self.context, v)
    elseif n == "border-top" then
        lib.xlFormatSetBorderTopA(self.context, v)
    elseif n == "border-bottom" then
        lib.xlFormatSetBorderBottomA(self.context, v)
    elseif n == "border-left-color" then
        lib.xlFormatSetBorderLeftColorA(self.context, v)
    elseif n == "border-right-color" then
        lib.xlFormatSetBorderRightColorA(self.context, v)
    elseif n == "border-top-color" then
        lib.xlFormatSetBorderTopColorA(self.context, v)
    elseif n == "border-bottom-color" then
        lib.xlFormatSetBorderBottomColorA(self.context, v)
    elseif n == "border-diagonal" then
        lib.xlFormatSetBorderDiagonalA(self.context, v)
    elseif n == "border-diagonal-style" then
        lib.xlFormatSetBorderDiagonalStyleA(self.context, v)
    elseif n == "border-diagonal-color" then
        lib.xlFormatSetBorderDiagonalColorA(self.context, v)
    elseif n == "fill-pattern" then
        lib.xlFormatSetFillPatternA(self.context, v)
    elseif n == "pattern-foreground-color" then
        lib.xlFormatSetPatternForegroundColorA(self.context, v)
    elseif n == "pattern-background-color" then
        lib.xlFormatSetPatternBackgroundColorA(self.context, v)
    elseif n == "locked" then
        lib.xlFormatSetLockedA(self.context, v)
    elseif n == "hidden" then
        lib.xlFormatSetHiddenA(self.context, v)
    else
        rawset(self, n, v)
    end
end

return format
