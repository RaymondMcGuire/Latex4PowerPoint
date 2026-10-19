---
name: latex4powerpoint
description: Add LaTeX formulas to PowerPoint slides the way the user's own Latex4PowerPoint add-in does, so the user can edit them later with the add-in. Use when a pptx made or changed with python-pptx needs a math formula, equation or symbol (数式, 公式), or when the user mentions Latex4PowerPoint, the formula add-in or 数学公式的插件. Also use to read the LaTeX code of the formulas in an existing deck.
---

# Latex4PowerPoint formulas

Latex4PowerPoint is the user's PowerPoint add-in (repository `D:\project\Latex4PowerPoint`, GitHub RaymondMcGuire/Latex4PowerPoint). It renders LaTeX code to a transparent PNG and stores the code in tags on the picture. The script `latex_eq.py` in this folder does the same from python-pptx. The user can then edit each formula in PowerPoint with the add-in.

## Setup

The script uses the add-in's own files, so the add-in must work on this PC first.

- `%APPDATA%\Latex4PowerPoint\settings.xml` gives the MiKTeX path. It exists after the user sets the path in the add-in options.
- `%APPDATA%\Latex4PowerPoint\LatexTemplate.txt` is the LaTeX template. It loads the user's notation package cgmathnotation.
- Python needs python-pptx, which brings Pillow and lxml.

To make sure that the setup works, run `python latex_eq.py` in this folder. It prints OK and the path of a test deck.

## Add a formula

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path.home() / ".claude" / "skills" / "latex4powerpoint"))
from latex_eq import add_latex

pic = add_latex(slide, r"$\trans{\mat{F}}\mat{F}=\mat{C}$", x, y, pt=20, color="2F363C")
```

- `x` and `y` are the top-left position in EMU (English Metric Units, 914400 per inch).
- `pt` is the font size that the formula shows at. Use the size of the body text near it. The width and height follow from it.
- `color` is a hex color or the add-in's "r,g,b" floats. Use the color of the nearby text, for example the deck's body text color.
- Write the code as a Python raw string. Use `$...$` for inline math and `equation*` or `aligned` for display math.
- After you add the picture, use `pic.width` and `pic.height` to center or align it.

If the code has a LaTeX error, `add_latex` raises an error with the end of the LaTeX log.

## Layout

Put a display formula on its own line or in its own box. A formula inside a sentence is a separate picture, so it does not move with the text. If you must put one inside a sentence, place it by coordinates, then render the slide and look at it.

## Notation (cgmathnotation)

Keep the user's own code as written. When you write a formula yourself, use the user's macros:

- `\mat{A}`, `\vect{x}`: bold matrix and vector. `\tensor{T}`: bold sans-serif tensor.
- `\trans{A}`: transpose with an upright sans-serif T. `\inv{A}`, `\invtrans{A}`: inverse and inverse transpose.
- `\trace`, `\diag`, `\rank`: upright operator names.
- `\pder{f}{x}`, `\tder{f}{x}`, `\pder[2]{f}{x}`: partial and total derivatives. `\diff`: upright d.
- `\norm{x}`, `\dotprod{a}{b}`, `\crossprod{a}{b}`, `\eu`, `\im`.

The full list is in `latex-sty\cgmathnotation_en.sty` in the repository.

## Reuse a formula from an existing deck

A formula made by the add-in keeps its code in the LATEXCODE tag of the picture. In the pptx zip, the tag is in `ppt/tags/tagN.xml`. The picture links to it through `<p:custDataLst><p:tags r:id="..."/>` in its `nvPr`. Newlines in the code are stored as the 4 characters `\r\n`.

## Test the result

- Render the slide with PowerPoint and look at it. In PowerShell, open the deck with `New-Object -ComObject PowerPoint.Application` and call `Slide.Export(path, "PNG", 1920, 1080)`.
- In the same COM session, `Shape.Tags.Item("LatexCode")` returns the code when the tags are correct.
- PowerShell variable names ignore case. Do not use `$s` for a slide and `$S` for a folder in the same script.

To edit a formula, the user selects it in PowerPoint and clicks the add-in's LaTeX button. The add-in renders the new code and keeps the current display size.
