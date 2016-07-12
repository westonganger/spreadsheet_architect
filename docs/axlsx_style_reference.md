# Basic Styling
- `:b`- Bold
- `:i`- Italic
- `:u`- Underline
- `:fg_color` - Text color
- `:bg_color` - Cell background color
- `alignment: {horizontal: true, vertical: true}` - Horizontal and/or Vertical center
- `font_name` - Font family
- `sz` - Font Size


# Spreadsheet Architect's More Sane Style Abstractions
I wanted a better and generic set of options for use between `xlsx` and `ods` so the following style aliases are available for use with Spreadsheet Architect:

- `:bold` == `:b`
- `:italic` == `:i`
- `:underline` == `:u`
- `:color` == `:fg_color`
- `:background_color` == `:bg_color`
- `:align` == `alignment: {horizontal: true, vertical: true}`
- `:font_size` == `sz`


# `number_format_code`
Here is the format for your `number_format_code` strings. 

This will output a dollar sign, comma's every three values, and minumum two decimal places: `\$#,##0.00`

This will output a nicely formatted date/time: `m/d/yyyy h:mm:ss AM/PM`


# `fmt_code`
