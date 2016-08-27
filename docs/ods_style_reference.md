```ruby
property :text, 
  'font-weight' => :bold, #options are :bold, :thin
  'font-size' => 12,
  'font-name' => 'Arial',
  'font-style' => 'italic',
  'text-underline-style' => 'solid', # solid, dashed, dotted, double
  'text-underline-type' => 'single',
  'text-line-through-style' => 'solid',
  align: true,
  color: "#000000"

property :cell, 
  'background-color' => "#DDDDDD",
  'wrap-option' => 'wrap',
  'vertical_align' => 'automatic',
  'border-top' => '0.75pt solid #999999',
  'border-bottom' => '0.75pt solid #999999',
  'border-left' => '0.75pt solid #999999',
  'border-right' => '0.75pt solid #999999',

property :column, 
  'column-width' => '4.0cm'

property :row, 
  'row-height' => '18pt',
  'use-optimal-row-height' => 'true'

property :table,
  'writing-mode' => 'lr-tb',
```
