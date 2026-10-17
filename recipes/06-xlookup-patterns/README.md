# XLOOKUP patterns that replace VLOOKUP headaches

Copy-paste lookups for the cases that trip people up: two criteria, whole columns at once, the latest match, tiered pricing.

- **Setup time:** 2 minutes
- **You'll need:** Google Sheets or Excel 365. Examples assume a `Products` tab with SKU in A, Size in B, Price in C.
- **Full page with copy buttons:** https://cogandcell.com/recipes/xlookup-patterns/

## The patterns

**Basic lookup with a friendly fallback**

```
=XLOOKUP(A2, Products!A:A, Products!C:C, "Not found")
```

**Return several columns at once (spills right)**

```
=XLOOKUP(A2, Products!A:A, Products!B:D)
```

**Two criteria (SKU + size), Google Sheets-safe**

```
=IFERROR(INDEX(FILTER(Products!C:C, Products!A:A=A2, Products!B:B=B2), 1), "Not found")
```

**Two criteria, Excel 365 style**

```
=XLOOKUP(1, (Products!A:A=A2)*(Products!B:B=B2), Products!C:C, "Not found")
```

**Most recent match (search bottom-up), e.g. last order date per customer**

```
=XLOOKUP(A2, Orders!B:B, Orders!A:A, "None", 0, -1)
```

**Tiered pricing: find the tier at or below a quantity (Tiers!A = min qty, B = unit price)**

```
=XLOOKUP(B2, Tiers!A:A, Tiers!B:B, , -1)
```

**Fill a whole column with one formula (Google Sheets)**

```
=ARRAYFORMULA(IF(A2:A="", , IFERROR(VLOOKUP(A2:A, Products!A:C, 3, FALSE), "Not found")))
```

**Works in every version of Excel and Sheets**

```
=IFERROR(INDEX(Products!C:C, MATCH(A2, Products!A:A, 0)), "Not found")
```

## Good to know

- Most “not found” errors are invisible differences: trailing spaces or numbers stored as text. Wrap the key in `TRIM()`, or see the [clean messy data](../clean-messy-data/) recipe.
- XLOOKUP defaults to an exact match, unlike VLOOKUP, which is one reason it's safer.
- In the tiered-pricing pattern, `-1` means “exact match, or the next smaller value”. Sort isn't required.

---
Need it set up for you? $49 fixed-price fix at https://cogandcell.com
