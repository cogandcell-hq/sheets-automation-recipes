# Split a full name into first and last name in Excel or Google Sheets

Formulas for every version of Excel and for Google Sheets, including middle names, 'Last, First' lists and extra spaces.

- **Setup time:** 2 minutes
- **You'll need:** A column of full names (examples assume A2). Excel 365 or Excel 2024 for TEXTBEFORE/TEXTAFTER; the classic versions work in any Excel and in Sheets.
- **Full page with copy buttons:** https://cogandcell.com/recipes/split-full-name/

## Excel 365 / 2024 (simplest)

**First name (text before the first space)**

```
=TEXTBEFORE(TRIM(A2), " ", 1, , , TRIM(A2))
```

**Last name (text after the last space)**

```
=TEXTAFTER(TRIM(A2), " ", -1, , , TRIM(A2))
```

**Middle name(s), blank if none**

```
=IFERROR(TEXTBEFORE(TEXTAFTER(TRIM(A2), " "), " ", -1), "")
```

**Both at once, spilling into two cells**

```
=HSTACK(TEXTBEFORE(TRIM(A2), " ", 1, , , TRIM(A2)), TEXTAFTER(TRIM(A2), " ", -1, , , TRIM(A2)))
```

The last argument (`TRIM(A2)`) is the “if not found” value, so a one-word name like *Cher* appears in both columns instead of showing #N/A.

## Any version of Excel (and Google Sheets)

**First name**

```
=LEFT(TRIM(A2), FIND(" ", TRIM(A2)&" ")-1)
```

**Last name (last word)**

```
=TRIM(RIGHT(SUBSTITUTE(TRIM(A2), " ", REPT(" ", 100)), 100))
```

The last-name trick swaps every space for 100 spaces, takes the right-most 100 characters (which can only contain the last word plus padding), and trims the padding.

## Google Sheets

**Split into one column per word**

```
=SPLIT(TRIM(A2), " ")
```

**First and last for a whole column at once**

```
=ARRAYFORMULA(IF(A2:A="", , {REGEXEXTRACT(TRIM(A2:A), "^\S+"), REGEXEXTRACT(TRIM(A2:A), "\S+$")}))
```

## Names written 'Last, First'

**First name (Excel 365)**

```
=TRIM(TEXTAFTER(A2, ","))
```

**Last name (any Excel or Sheets)**

```
=TRIM(LEFT(A2, FIND(",", A2&",")-1))
```

## Good to know

- No formulas needed for a one-off: select the column, then **Data → Text to Columns** (Excel) or **Data → Split text to columns** (Sheets). Excel's **Flash Fill** (Ctrl+E) also works after you type two examples.
- Surnames with spaces (*van der Berg*, *de la Cruz*) and suffixes (*Jr.*) can't be split reliably by any formula. Review those by hand, or have us build a rule list for your data.
- Non-breaking spaces from web copy defeat TRIM. Clean them first with `SUBSTITUTE(A2, CHAR(160), " ")`; see [clean messy data](../clean-messy-data/).

---
Want it done for you? Fixed price, free quote: see pricing at https://cogandcell.com/pricing
