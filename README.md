# spreadsheet-builder

A simple spreadsheet builder tool.

This template language is designed to generate XLSX spreadsheets effectively. The underlying library is **rust_xlsxwriter**.

---

## Template structure

The template has two main sections:

1. **Format definitions**  
2. **Sheet & content definitions**

---

## 1. Formats

You can define reusable cell formats:

```rust
:header {
  border("thin"),
  border_bottom_color("#000000"),
  background_color("#eeeeee")
}
```

Format identifiers:

- Must start with `:`  
- Can contain: a-z, A-Z, 0-9, _

**Supported format properties:**

- `bold`  
- `italic`  
- `underline`  
- `strikethrough`  
- `super`  
- `sub`  
- `num("")`  
- `align("left")`  // left, right, center, verticalcenter, top, bottom  
- `indent(1)`  
- `font_name("")`  
- `font_size(12)`  
- `color("#")`  
- `background_color("#")`  
- `border("thin")` // medium, dashed, dotted, thick, double, hair, etc.  
- `border_top(...)`  
- `border_bottom(...)`  
- `border_left(...)`  
- `border_right(...)`  
- `border_color("#")`  
- `border_top_color(...)`  
- `border_bottom_color(...)`  
- `border_left_color(...)`  
- `border_right_color(...)`

**Note:** Dates are stored as numbers in Excel → use `num()` formatting for display.

---

## 2. Sheets & layout

```rust
sheet("Data")
row(0, pixels(75))
col(0, 0, pixels(186))
```

- `sheet()` → starts a new worksheet  
- `row()` → sets row height  
- `col()` → sets column width  

---

## Cells & content

Example:

```rust
[
  img("images/alc-logo.png", :border, embed),
  str($report_data.title, :maintitle, colspan(8)),
  str($report_data.username, :right, colspan(2))
]
```

Supported cell types:

- `str()` → string  
- `num()` → number  
- `date()` → ISO → Excel date  
- `img()` → image  

**Image modes:**

- `embed` → fits inside cell  
- `insert` → may overflow cells  

**Merging:**

- `colspan(n)`  
- `rowspan(n)`  

---

## Variables & loops

You can iterate over arrays:

```rust
for $prize in $prize_levels {
  [
    str($prize.prize_level_no),
    str($prize.description),
    num($prize.amount)
  ]
}
```

---

## Horizontal for loops

Generate multiple cells within a single row:

```rust
[ str("Start"), for $val in $arr { str($val) }, str("End") ]
```

Expands into:

```
Start | val1 | val2 | val3 | End
```

**Multiple loops in one row:**

```rust
[
  for $val in $arr { str($val) },
  for $other in $otherArr { str($other) }
]
```

- Each loop expands independently  
- Final row concatenates all generated cells  

**Notes:**

- Works only at row level  
- Ideal for dynamic column generation  
- Fully compatible with other expressions  

---

## header() – Dynamic merged headers

Designed for dynamic column layouts (especially XLSX).

**Input format:**

```rust
$headers = [
  ["Name", 2],
  ["Score", 3],
  ["Notes", 1]
]
```

- `String` → label  
- `usize` → number of columns to span  

**Usage:**

```rust
header($headers, :gray)
```

**XLSX behavior:**

- Cells are merged horizontally  
- Format applied to merged region  

**CSV behavior:**

```
Name,,Score,,,Notes
```

- No merge support → padded with empty cells  
- Keeps column count consistent  

**Inline usage:**

```rust
[ str("hello"), header($headers, :gray) ]
```

---

## Conditional blocks (if)

```rust
if $show_header {
  [
    str("Header1"),
    str("Header2")
  ]
}
```

**Optional else:**

```rust
if $cond {
  [ ... ]
} else {
  [ ... ]
}
```

Else is optional.

---

## Conditions (extended)

**Supported operators:**

```
==  !=  <  >  <=  >=
```

**Supported types:**

- string  
- integer  
- float  

**Examples:**

```rust
if $myinteger == 3
if $myinteger < 5
if $myfloat == 5.0
if $mystring == "abc"
if $mystring < $mystring2   // lexicographical
```

**Important:**

- Type mismatch → evaluates to false  

---

## Cursor control

```rust
anchor(@top)
move(@top, 0, 3)
move(0, 3)
cr
```

- `anchor()` → save position  
- `move()` → jump relative or absolute  
- `cr` → move to beginning of row  

---

## Autofit

```rust
autofit
```

- Attempts automatic column sizing  
- ⚠️ Not always reliable → prefer explicit widths  

---

## Notes & limitations

**CSV:**

- No merge support  
- No styling  

- Each row must have consistent number of columns  
- Horizontal expansion (loops, header) helps maintain this dynamically  

---

## Full example template

```rust
:header {
  bold,
  background_color("#cccccc"),
  align("center")
}

:gray {
  background_color("#eeeeee")
}

sheet("Report")
row(0, pixels(30))

$headers = [
  ["Name", 2],
  ["Score", 3],
  ["Notes", 1]
]

[ header($headers, :gray) ]

for $student in $students {
  [
    str($student.first_name),
    str($student.last_name),
    for $score in $student.scores { num($score) },
    str($student.note)
  ]
}

if $show_summary {
  [
    str("Average", :header),
    "",
    for $col_avg in $averages { num($col_avg) },
    ""
  ]
}
```

- Dynamic headers (`header`)  
- Horizontal loop for scores  
- Conditional summary row  

---

## Best Practices

- Always define formats for merged/important cells    
- Ensure consistent column count per row → avoids broken CSV/XLSX  
- Use anchors and move for complex layouts → easier to maintain  
- Prefer explicit column widths over autofit → more predictable output  
- Use `header()` in CSV cautiously → merges are simulated via empty cells  
- Type consistency in conditions → mismatched types evaluate to false  
- Combine loops and if for compact dynamic tables
