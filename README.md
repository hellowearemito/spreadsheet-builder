# spreadsheet-builder
A simple spreadsheet builder tool

This template language is designed to generate xlsx spreadsheets effectively. 
The underlying library is the rust_xlsxwriter library 

The template has two sections. First, you define cell formats like:

```
:header {
  border("thin"),
  border_bottom_color("#000000"),
  background_color("#eeeeee")
}  
```

Format identifiers always begin with : followed by US-ASCII letters, digits or underscore.

Handled formats:

- bold
- italic
- underline
- strikethrough
- super
- sub
- num(“<format>”) 
- align(“left”) - or right, center, verticalcenter, top, bottom
- indent(1)
- font\_name(“<name>”)
- font\_size(12)
- color(“#<hexa>”)
- background\_color(“#<hexa>”)
- border(“thin”)   - or medium, dashed, dotted, thick, double, hair, medium\_dashed, dash\_dot, medium\_dash\_dot, dash\_dot\_dot, medium\_dash\_dot\_dot, slant\_dash\_dot
- border\_top(..)
- border\_bottom(..)
- border\_left(..)
- border\_right(..)
- border\_color(“#<hexa>”)
- border\_top\_color(..)
- border\_bottom\_color(..)
- border\_left\_color(..)
- border\_right\_color(..)

Dates are numbers in Excel, so their formatting is handled by the num() format as well.

In the second section, you define the sheets:

```
sheet("Data")
row(0, pixels(75))
col(0, 0, pixels(186))
```

The sheet() starts a new WorkSheet. The row() sets the height of a row (either in pixels or in chars).

The col() sets the width for a range of columns (either in pixels or in chars).

This snippet generates two rows and cells into them (starting at the current cursor position):


```
[ 
  img("images/alc-logo.png", :border, embed), 
  str($report_data.title, :maintitle, colspan(8)), 
  str($report_data.username, :right, colspan(2)) 
]
[ 
  str(""), 
  str($report_data.jurisdiction, :center, colspan(8)), 
  str($report_data.module, :right, colspan(2)) 
]
```

The img() type inserts an image into the cell, optionally applies a format on the cell. The image placement may be embed or insert. The first one fits the image inside the cell, the second one allows it to overflow multiple cells.

The str() creates a cell with string content, the num() with numeric content and the date() with date content. Dates are actually numbers in Excel, the date() call converts an ISO 8601 timestamp string into an Excel number.

The colspan() and rowspan() modifiers set the cell merging properties.

You can use the passed variables like this:

```
[ str("Game:", :header), str($instant_game.game) ]
```

or like this:

```
for $prize in $prize_levels {
  [
    str($prize.prize_level_no),
    str($prize.description),
    str($prize.prize_code),
    num($prize.amount),
    num($prize.free_ticket, :int),
    str($prize.merch_prize),
    num($prize.prize_value),
    num($prize.wins, :int),
    num($prize.percent_of_sale)
  ]
}
```

The expression language may be improved in the feature if necessary.

For movement of the cursor you can use two statements: anchor and move

```
anchor(@top)
```

remembers the current position, named as @top. Anchor names always start with a @ follow by US-ASCII letters, digits or underscore.


```
move(@top, 0, 3)
```

moves the cursor to a position 0 rows below @top and 3 columns right.

You can move relative to the current position too:


```
move(0, 3)
```

To return to the beginning of a row, use

```
cr
```

There is an 

```
autofit
```

command to automatically set the width of the columns based on their content but this is not 100% reliable, setting explicit column width is preferred.


