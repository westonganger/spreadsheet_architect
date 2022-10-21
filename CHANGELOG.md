CHANGELOG
---------

- **5.0.0 - Unreleased** - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v4.2.0...master)
  - Update to caxlsx v3.3.0+ which now contains the axlsx_styler code, so we drop the dependency on axlsx_styler
  - [#38](https://github.com/westonganger/spreadsheet_architect/pull/38) - **Breaking Change**  - Add `escape_formulas` option for xlsx spreadsheets. This is a breaking change because we default to `escape_formulas: true` whereas before there was no formula escaping at all. The reasoning for this breaking change is that creating spreadsheets where many of the fields contain direct user input are a large majority compared to use cases that involve formulas.
  - [#39](https://github.com/westonganger/spreadsheet_architect/pull/39) - Add option `use_zero_based_row_index: true` (Default `false`) which allows you to use zero-based row indexes instead of the default 1-based row indexes. Recomended to set this option for the whole project. The original reason it was designed to be 1-based is because spreadsheet row numbers literally start with 1. However this tends to be unituitive for the developer because columns use zero based indexes because they use letter-based notation instead.
  - [#40](https://github.com/westonganger/spreadsheet_architect/pull/40) - Improve argument handling for freeze option and add support for all Axlsx supported options for panes using the `:freeze` hash. See test case for example (./test/unit/xlsx_freeze_test.rb)
  - [#42](https://github.com/westonganger/spreadsheet_architect/pull/42) - Improve exceptions and messages regarding invalid ranges
  - [#45](https://github.com/westonganger/spreadsheet_architect/pull/45) - For `to_xlsx`, dont add empty header row when `header: true`
  - [#44](https://github.com/westonganger/spreadsheet_architect/pull/44) - Add support for hyperlinks in XLSX and ODS
  - [#49](https://github.com/westonganger/spreadsheet_architect/pulls/49) - Extracted some ODS methods from `SpreadsheetArchitect::Utils` to `SpreadsheetArchitect::Utils::ODS`

- **4.2.0** - May 27, 2021 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v4.1.0...v4.2.0)
  - Add option `:skip_defaults` which removes the defaults and default styles. Particularily useful for heavily customized spreadsheets where the default styles get in the way.
  - Fix bug where styles werent being un-applied when using the `false` value.
  - Add style aliases for `:valign` and `:wrap_text`
  - Fix error with `headers: false`, previously had to use `headers: []`

- **4.1.0** - Nov 20, 2020 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v4.0.1...v4.1.0)
  - Raise ArgumentError when invalid option names are given

- **4.0.1** - Nov 20, 2020 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v4.0.0...v4.0.1)
  - Fix bug with `headers: false` where a blank header row is still added 
  - Fix Bug for older version of `caxlsx` v2.0.2

- **4.0.0** - Mar 3, 2020 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v3.3.1...v4.0.0)
  - Switch to the `caxlsx` gem (Community Axlsx) from the legacy unmaintained `axlsx` gem. Axlsx has had a long history of being poorly maintained so this community gem improves the situation.
  - Require Ruby 2.3+
  - Ensure all options using Hash are automatically converted to symbol only hashes
  - Add XLSX option `:freeze` to freeze custom sections of your spreadsheet
  - Add XLSX option `:freeze_headers` to freeze the headers of your spreadsheet
  - Remove old Axlsx patch for column width
  - Backport new code for `string_width` calculations to Axlsx 3.0.1 and below.

- **3.3.1** - Dec 2, 2019 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v3.3.0...v3.3.1)
  - [Issue #30](https://github.com/westonganger/spreadsheet_architect/issues/30) - Fix duplicate constant warning for XLSX_COLUMN_TYPES

- **3.3.0** - Nov 28, 2019 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v3.2.1...v3.3.0)
  - Fix `:borders` option, was broken in v3.2.1
  - Fix bug when passing `false` to `:headers` option
  - Raise error when unsupported column type is passed
  - Remove claimed support for `:currency` and `:percent` for ODS spreadsheets as they were not working. PR Wanted.

- **3.2.1** - April 10, 2019 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v3.2.0...v3.2.1)
  - Fix bug when using `column_style` option with `include_header: true` & letter based column numbering

- **3.2.0** - September 14, 2018 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v3.1.0...v3.2.0)
  - Change implementation of `:column_styles` option to utilize `axlsx_styler` instead of the built-in axlsx `col_style` method. The reason for the switch is that `col_style` would overwrite all previously set styles. `axlsx_styler` already has the ability to add onto existing styles and is what is currently utilized by `range_styles`.
  - Date / Time formatting is now set per cell instead of on the entire column.
  - Default Date formatting for `xlsx` changed from `m/d/yyyy` to `yyyy-mm-dd`
  - Default Time/DateTime formatting for `xlsx` changed from `yyyy/m/d h:mm AM/PM` to `yyyy-mm-dd h:mm AM/PM`
  - Fix bug where the ActionController::Renderer `:filename` option was ignored when an AR::Relation passed directly to the renderer without first calling `to_#{format}`

- **3.1.0** - August 19, 2018 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v3.0.0...v3.1.0)
  - Add new option `:conditional_row_styles` to `to_xlsx`.
  - Add ability to pass an alternative method name as a Symbol/String to the `:spreadsheet_columns` option.
  - Replace all usage of the legacy method `instance_eval` with the proper method `send`.
  - [#23](https://github.com/westonganger/spreadsheet_architect/issues/23#issuecomment-412803761) - Fix bug where custom `columns_widths` in xlsx spreadsheets might not get set correctly.
  - All exceptions now inherit from the appropriate ruby core exception classes
  - `SpreadsheetArchitect::Exceptions::InvalidOptionError` renamed to `SpreadsheetArchitect::Exceptions::OptionTypeError`

- **3.0.0** - July 6, 2018 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v2.1.2...v3.0.0)
  - [#16](https://github.com/westonganger/spreadsheet_architect/issues/16) - Add ability to pass :instances option to SpreadsheetArchitect class methods
  - [#16](https://github.com/westonganger/spreadsheet_architect/issues/16) - Remove Plain Ruby syntax `Post.to_xlsx(instances: posts_array)` in favor of `SpreadsheetArchitect.to_xlsx(instance: posts_array)`. However, it may still work at this time if configured correctly.
  - Fix project-wide and model-level defaults before only `header_style`, `row_style`, & `sheet_name` were being utilized.
  - When using on an ActiveRecord class and `spreadsheet_columns` is not defined, it now defaults to the classes `column_names` only. Previously it would use `column_names` and then remove the following columns `['id', 'created_at', 'updated_at', 'deleted_at']`
  - XLSX column ranges now also accept letters. For example: `{columns: ('C'..'E')}`
  - `:column_types` now considers types defined in `spreadsheet_columns` and class/project-wide defaults. Before it was incorrectly ignored.
  - Passing the `spreadsheet_columns` options now only accepts lambda/proc
  - More type checking and Option types are now being type checked. Option types were supposed to be properly type checked but due to a bug were being skipped.
  - Utilize ActiveSupport `pluralize`, if available, for default sheet names for class-based spreadsheets
  - Renamed `BadRangeError` to `InvalidRangeError`
  - Renamed `IncorrectTypeError` to `InvalidTypeError`
  - Remove all Rails generators `spreadsheet_architect:add_default_options`. No need since its just as easy to copy from the README
  - Major overhaul of test suite, add a ton more tests, for DRYness use resursion for tests when appropriate
  - Use appraisal to test various `axlsx` versions

- **2.1.2** - July 6, 2018 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v2.1.1...v2.1.2)
  - Fix bug where everything was underlined by default in Excel (LibreOffice was working correctly). For some reason, `false` in `:u` or `:underline` was incorrectly being treated as `true` but only within Excel. Now anytime `false` is encountered for either `:u` or `:underline` it is now converted to `nil`
  - Fix bug where empty xlsx spreadsheets were corrupt when trying to open with Excel (LibreOffice was working correctly). This only occured when containing no headers and empty `:data` option which resulted in a package with no sheets.

- **2.1.1** - July 4, 2018 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v2.1.0...v2.1.1)
  - [#18](https://github.com/westonganger/spreadsheet_architect/pull/18) - Fix controller bug when using an non-ActiveRecord ORM only within Rails

- **2.1.0** - June 20, 2018 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v2.0.2...v2.1.0)
  - [#15](https://github.com/westonganger/spreadsheet_architect/pull/15) - Improved the method symbolize_keys. This method did not work properly for nested objects.
  - [PR #15](https://github.com/westonganger/spreadsheet_architect/pull/15) - Added the ability to pass `:text_wrap` option within the `:alignment` style
  - Make axlsx styles higher precendence over Spreadsheet Architect style aliases
  - Use `prepend` monkey patches in Ruby 2+ to avoid annoying overwrite warnings when using old `define_method` monkey patches
  - Due to [RODF bug](https://github.com/thiagoarrais/rodf/issues/19) convert all Date and Time cells to String in ODS spreadsheets
  - Improve test suite
  - Dont test against Ruby versions that Rails no longer supports. Gem code should remain compatible with Ruby 1.9.3.

- **2.0.2** - July 14 2017 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v2.0.1...v2.0.2)
  - Fix bug with range styles rows option not counting headers
  - Fix bug with range styles rows :all option

- **2.0.1** - February 16 2017 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v2.0.0...v2.0.1)
  - Fix bug where `SpreadsheetArchitect.default_options` and `SPREADSHEET_OPTIONS` were being overwritten
  - Fix bug where col_styles ignored previous styles on header when using `include_header` option
  - Errors now try to provide which value is the cause

- **2.0.0** - January 28 2017 - [View Diff](https://github.com/westonganger/spreadsheet_architect/compare/v1.4.8...v2.0.0)
  - Add to xlsx: `merges`, `column_styles`, `range_styles`, `borders`, `column_widths` multi-row headers, date/time default format_code
  - Add `column_types` option for xlsx and ods
  - Add ability to make multi-sheet spreadsheets in XLSX & ODS
  - Adds `axlsx_styler` gem dependency
  - Add Examples
  - Add Axlsx Style Reference
  - Refractor into smaller files

- **1.4.8** - December 6 2016
  - Lock `rodf` gem to v0.3.7 for last v1 version of this gem

- **1.4.7** - November 7 2016
  - Fix method arguments for `to_rodf_spreadsheet` method

- **1.4.6** - May 16 2016
  - Fix hash syntax for support of ruby v2.1 and below

- **1.4.5** - May 4 2016
  - Bug fixes

- **1.4.4** - May 3 2016
  - Add Ability to add format_code to all numbers body rows

- **1.4.3** - May 3 2016
  - Bug fixes

- **1.4.2** - May 3 2016
  - Add to_axlsx_package, to_rodf_spreadsheet methods for the item to be further manipulated. Ex. axlsx_styler

- **1.4.1** - May 2 2016
  - Add rails generator for project defaults initializer

- **1.4.0** - April 29 2016
  - Add to_xlsx, to_ods, & to_csv to SpreadsheetArchitect model for direct calling by passing in cell data

- **1.3.0** - April 21 2016
  - Add ability to create class/model and project option defaults

- **1.2.5** - March 25 2016
  - Fix each_with_index bug

- **1.2.4** - March 24 2016
  - Fix cell type logic for symbol methods

- **1.2.3** - March 20 2016
  - Fix cell type logic

- **1.2.2** - March 19 2016
  - Make cell type numeric if value is numeric

- **1.2.1** - March 13 2016
  - Better error reporting
  - Fix for Plain ruby models

- **1.2.0** - March 10 2016
  - Fix Bug: first row data repeated for all records on custom values

- **1.1.0** - March 3 2016
  - Breaking Change - Move spreadsheet_columns method from the class to the instance
  - Fix Bug: remove default underline on cells

- **1.0.4** - March 1 2016
  - Extract helper methods to seperate module
  - Improve readme

- **1.0.3** - March 1 2016
  - Fix/Improve renderers
  - Fix header default background color
  - Fix default columns

- **1.0.2** - February 26 2016
  - Enhance Style options

- **1.0.1** - February 26 2016
  - Fix bug in renderers

- **1.0.0** - February 26 2016
  - Gem Initial Release
