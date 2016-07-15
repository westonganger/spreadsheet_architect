# Basic Styling
- `b` (Boolean) - Bold
- `i` (Boolean) - Italic
- `u` (Boolean) - Underline
- `fg_color` (String) - Text Color - Ex: `000000`
- `bg_color` (String) - Cell background color - Ex: `CCCCCC`
- `alignment: {horizontal: true, vertical: true}` (Hash) - Horizontal and/or Vertical center
- `strike` (Boolean) — Indicates if the text should be rendered with a strikethrough
- `outline` (Boolean) — Indicates if the text should be rendered with a shadow
- `sz` (Integer) - Font Size
- `font_name` (String) - Font Name - Ex. `Arial`
- `family` (Integer) — The font family to use. `1` = Roman (Default), `2` = Swiss, `3` = Modern, `4` = Script, `5` = Decorative
- `charset` (Integer) — The character set to use. Axlsx documentation says this setting is ignored most of the time.
- `type` (Symbol) - Type of the cell. Options are: `:xf` (Default) or `:dxf`
- `border: {style: :thin, color: "000000", edges: [:top, :bottom, :left, :right]}` (Hash) - Borders support style, color and edges options. Available styles for the border are: `:none, :thin, :medium, :dashed, :dotted, :thick, :double, :hair, :mediumDashed, :dashDot, :mediumDashDot, :dashDotDot, :mediumDashDotDot, :slantDashDot`
- hidden (Boolean) — Indicates if the cell should be hidden
- locked (Boolean) — Indicates if the cell should be locked
- format_code (String) — See string formats in `num_fmt` options below
- num_fmt (Integer) — I find it much more preferable to write the `format_code` manually instead of using this option.
  - 1 = '0' 
  - 2 = '0.00' 
  - 3 = '#,##0'
  - 4 = '#,##0.00'
  - 5 = '$#,##0_);($#,##0)'
  - 6 = '$#,##0_);Red'
  - 7 = $#,##0.00_);($#,##0.00)'
  - 8 = '$#,##0.00_);Red'
  - 9 = '0%'
  - 10 = '0.00%' 
  - 11 = '0.00E+00'
  - 12 = '# ?/?'
  - 13 = '# ??/??'
  - 14 = 'm/d/yyyy'
  - 15 = 'd-mmm-yy'
  - 16 = 'd-mmm'
  - 17 = 'mmm-yy'
  - 18 = 'h:mm AM/PM'
  - 19 = 'h:mm:ss AM/PM'
  - 20 = 'h:mm'
  - 21 = 'h:mm:ss'
  - 22 = 'm/d/yyyy h:mm'
  - 37 = '#,##0_);(#,##0) 38 #,##0_);Red'
  - 39 = '#,##0.00_);(#,##0.00)'
  - 40 = '#,##0.00_);Red'
  - 45 = 'mm:ss'
  - 46 = '[h]:mm:ss'
  - 47 = 'mm:ss.0'
  - 48 = '##0.0E+0'
  - 49 = '@'

# Some style aliases provided by Spreadsheet Architect
I wanted a better and generic set of options for use between `xlsx` and `ods` so the following style aliases are available for use with Spreadsheet Architect:

- `bold` is the same as `b`
- `italic` is the same as `i`
- `underline` is the same as `u`
- `color` is the same as `fg_color`
- `background_color` is the same as `bg_color`
- `align` is the same as `alignment: {horizontal: true, vertical: false}`
- `font_size` is the same as `sz`


# `format_code`
This will output a dollar sign, comma's every three values, and minumum two decimal places: 
`\$#,##0.00`

This will output a nicely formatted date/time: 
`m/d/yyyy h:mm:ss AM/PM`

See the options for `num_fmt` above to see more examples of how this can be used.
