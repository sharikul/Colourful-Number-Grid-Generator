---
Original project name: Colourful Number Grid Generator
Current project name: Colourful Number Grid Generator
Start date: 2026-09-05
Current status: Complete
Initial completion date: 2026-09-09
Language: Python
---
# Colourful Number Grid Generator
* **Project started**: Saturday, 05 September 2026
* **Project completed**: Wednesday, 09 September 2026

## Background
A playful script to generate a grid of numbers starting from 1 to the number supplied as the argument (default is 100), which is exported to Excel. Square numbers are highlighted in red, and numbers that cross the grid are coloured in orange, to produce a colourful grid that when zoomed out looks similar to the Union Jack flag.

### Working files
* **CNGG.py**: script file
* **griddy.json**: list of input numbers that generate a clean ('griddy') grid, containing an equal number of rows and columns, with a cross grid that go from beginning to end and vice versa. Populated by running `cngg()` in long while loops with different input numbers that were multiplied in every loop.

### Python libraries used
#### Pre-installed:
---
* pathlib (for function `Path`)
* json
#### Dependencies (requiring installation if not already installed):
---
* xlsxwriter (to export to Excel)

## Functions/Classes

### `cngg`
```python
def cngg(inp:int = 100, ind:bool = False)
```
#### Args

| Parameter | Type | Notes |
| --- | --- | --- |
| `inp` | Mandatory | Set to 100 if left empty  |
| `ind` | Mandatory | Set to False if left empty. If `True`, a separate Excel file is created for a grid, with the input number included in the file name|
    