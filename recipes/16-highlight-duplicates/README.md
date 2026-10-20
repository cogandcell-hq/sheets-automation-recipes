# Highlight duplicates in Google Sheets or Excel with conditional formatting

Custom formulas that color duplicate values, whole duplicate rows, or only the second and later copies, so you can review before deleting anything.

- **Setup time:** 3 minutes
- **You'll need:** Google Sheets or Excel. Examples assume the values to check are in column B with a header in row 1.
- **Full page with copy buttons:** https://cogandcell.com/recipes/highlight-duplicates/

## Google Sheets steps

1. Select the range to color, e.g. `B2:B` (or `A2:Z` to color whole rows).
2. **Format → Conditional formatting**. Under *Format cells if*, choose **Custom formula is**.
3. Paste one of the formulas below, pick a fill color, and click **Done**.

**Every copy of a duplicate value (apply to B2:B)**

```
=COUNTIF($B$2:$B, $B2)>1
```

**Whole row, based on column B (apply to A2:Z)**

```
=COUNTIF($B$2:$B, $B2)>1
```

**Only the 2nd, 3rd... copies, so the first stays unhighlighted**

```
=COUNTIF($B$2:$B2, $B2)>1
```

**Duplicates across two columns together (first name in A + last name in B)**

```
=COUNTIFS($A$2:$A, $A2, $B$2:$B, $B2)>1
```

**Ignore blank cells**

```
=AND($B2<>"", COUNTIF($B$2:$B, $B2)>1)
```

## Excel steps

For simple cases use **Home → Conditional Formatting → Highlight Cells Rules → Duplicate Values**. For whole rows or “2nd copy only”, choose **New Rule → Use a formula to determine which cells to format**, select e.g. `A2:Z1000` first, and use the same formulas with a fixed end row:

**Excel: whole row, based on column B**

```
=COUNTIF($B$2:$B$1000, $B2)>1
```

## Good to know

- The dollar signs matter: `$B2` locks the column but lets the row move, so each row checks its own value.
- COUNTIF ignores upper/lower case, so `ANN@X.COM` and `ann@x.com` are flagged as duplicates. Trailing spaces are *not* ignored; clean them first.
- Ready to delete them? See [remove duplicate rows](../remove-duplicate-rows/) or [keep the latest entry](../remove-duplicates-keep-latest/).

---
Want it done for you? Fixed price, free quote: see pricing at https://cogandcell.com/pricing
