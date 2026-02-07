# Custom Excel LAMBDA Functions

> Paste each formula into **Formulas → Name Manager → New**.
> Set the **Name**, paste the **Formula** into **Refers to**, and add the **Comment** as the description.

---

## Text / Data Cleaning

### TEXTBETWEEN

| Field         | Value                                                                                                                                                      |
| ------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Name**      | `TEXTBETWEEN`                                                                                                                                              |
| **Comment**   | Extracts text between two delimiters. Returns empty string if not found.                                                                                   |
| **Refers to** | `=LAMBDA(text, start_delim, end_delim, LET(s, FIND(start_delim, text) + LEN(start_delim), e, FIND(end_delim, text, s), IFERROR(MID(text, s, e - s), "")))` |

#### Usage

```excel
=TEXTBETWEEN(A1, "(", ")")
=TEXTBETWEEN("Hello [World]", "[", "]")   → World
```

---

### EXTRACTNUMBERS

| Field         | Value                                                                                                                     |
| ------------- | ------------------------------------------------------------------------------------------------------------------------- |
| **Name**      | `EXTRACTNUMBERS`                                                                                                          |
| **Comment**   | Returns only numeric characters (0-9) from a text string.                                                                 |
| **Refers to** | `=LAMBDA(text, LET(chars, MID(text, SEQUENCE(LEN(text)), 1), nums, IF(ISNUMBER(VALUE(chars)), chars, ""), CONCAT(nums)))` |

#### Usage

```excel
=EXTRACTNUMBERS("INV-2026-0042")   → 20260042
```

---

### EXTRACTLETTERS

| Field         | Value                                                                                                                                                             |
| ------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Name**      | `EXTRACTLETTERS`                                                                                                                                                  |
| **Comment**   | Returns only alphabetic characters (A-Z, a-z) from a text string.                                                                                                 |
| **Refers to** | `=LAMBDA(text, LET(chars, MID(text, SEQUENCE(LEN(text)), 1), codes, CODE(UPPER(chars)), letters, IF((codes >= 65) * (codes <= 90), chars, ""), CONCAT(letters)))` |

#### Usage

```excel
=EXTRACTLETTERS("Order#12345-AB")   → OrderAB
```

---

### COUNTWORDS

| Field         | Value                                                                                             |
| ------------- | ------------------------------------------------------------------------------------------------- |
| **Name**      | `COUNTWORDS`                                                                                      |
| **Comment**   | Counts the number of words in a text string (space-separated). Returns 0 for blank cells.         |
| **Refers to** | `=LAMBDA(text, IF(ISBLANK(text), 0, LEN(TRIM(text)) - LEN(SUBSTITUTE(TRIM(text), " ", "")) + 1))` |

#### Usage

```excel
=COUNTWORDS("Hello beautiful world")   → 3
```

---

## How to Install

1. Open your workbook (or a **Personal Macro Workbook** for global access).
2. Go to **Formulas → Name Manager** (or press `Ctrl+F3`).
3. Click **New**.
4. Enter the **Name** exactly as shown above.
5. Paste the **Refers to** formula.
6. Add the **Comment** as the description.
7. Set **Scope** to **Workbook**.
8. Click **OK**.

> **Tip:** To make LAMBDAs available in every workbook, define them in your **Personal.xlsb** file.
