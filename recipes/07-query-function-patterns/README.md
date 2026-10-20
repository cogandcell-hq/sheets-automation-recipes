# QUERY function patterns: totals by customer, by month, top 5 and more

Build live summaries with one formula: filter, group, pivot and sort, without pivot tables.

- **Setup time:** 5 minutes
- **You'll need:** Google Sheets. Examples assume a `Sales` tab: A Date, B Customer, C Product, D Qty, E Amount, with headers in row 1.
- **Full page with copy buttons:** https://cogandcell.com/recipes/query-function-patterns/

## The patterns

**Filter and sort: orders over $100, biggest first**

```
=QUERY(Sales!A:E, "select A, B, E where E > 100 order by E desc", 1)
```

**Revenue by customer**

```
=QUERY(Sales!A:E, "select B, sum(E) where B is not null group by B order by sum(E) desc label sum(E) 'Revenue'", 1)
```

**Top 5 products by quantity**

```
=QUERY(Sales!A:E, "select C, sum(D) where C is not null group by C order by sum(D) desc limit 5 label sum(D) 'Units'", 1)
```

**Revenue by month (month() is 0-based, hence +1)**

```
=QUERY(Sales!A:E, "select year(A), month(A)+1, sum(E) where A is not null group by year(A), month(A)+1 label year(A) 'Year', month(A)+1 'Month', sum(E) 'Revenue'", 1)
```

**Between two dates typed in G1 and G2**

```
=QUERY(Sales!A:E, "select * where A >= date '"&TEXT(G1, "yyyy-mm-dd")&"' and A <= date '"&TEXT(G2, "yyyy-mm-dd")&"'", 1)
```

**Search box: product contains whatever is typed in G3**

```
=QUERY(Sales!A:E, "select * where lower(C) contains '"&LOWER(G3)&"'", 1)
```

**Pivot: customers down, products across**

```
=QUERY(Sales!A:E, "select B, sum(E) where B is not null group by B pivot C", 1)
```

## Good to know

- The last argument (`1`) is the number of header rows. Getting it wrong is the most common cause of odd results.
- QUERY uses the majority data type in each column; minority values (e.g. a few text entries in a number column) are treated as empty. Keep columns consistent.
- Columns are letters when querying a range on a sheet, but `Col1, Col2…` when querying an array such as `{...}` or IMPORTRANGE.
- Dates in a query string must use the `date 'yyyy-mm-dd'` format, which is why the TEXT() wrapper is there.

---
Want it done for you? Fixed price, free quote: see pricing at https://cogandcell.com/pricing
