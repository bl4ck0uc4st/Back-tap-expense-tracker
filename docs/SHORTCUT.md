# Shortcut reference

A compact reference for the iOS side. Full build instructions are in
[SETUP.md](SETUP.md).

## Action list

| # | Action | Configuration |
|---|---|---|
| 1 | Ask for Input | Number · "Enter Amount" |
| 2 | Set Variable | `Amt` ← Ask for Input |
| 3 | Choose from Menu | "Income or Expense?" |
| 3a | ‣ Income branch | Text `Income` → Set Variable `Typ` |
| 3b | ‣ ‣ Choose from Menu | "Category" · 4 income branches |
| 3c | ‣ Expense branch | Text `Expense` → Set Variable `Typ` |
| 3d | ‣ ‣ Choose from Menu | "Category" · 9 expense branches |
| 4 | Ask for Input | Date · "Date of expense?" |
| 5 | Format Date | Custom · `yyyy-MM-dd` |
| 6 | Set Variable | `Dt` ← Formatted Date |
| 7 | Ask for Input | Text · "Description" |
| 8 | URL Encode | ← Ask for Input |
| 9 | Set Variable | `Desc` ← URL Encoded Text |
| 10 | Text | the request URL |
| 11 | Get Contents of URL | ← Text · GET · no headers |
| 12 | Show Notification | confirmation |

Every category branch is the same two actions: a **Text** holding the value,
then a **Set Variable** named `Cat` pointing at it.

## Variables

| Name | Holds | Encoded |
|---|---|---|
| `Amt` | amount | no |
| `Typ` | `Income` or `Expense` | no |
| `Cat` | category name | no — keep it single-word |
| `Dt` | `yyyy-MM-dd` | no |
| `Desc` | free text | yes |

## Request URL

```
https://script.google.com/macros/s/YOUR_ID/exec?am=[Amt]&ty=[Typ]&ct=[Cat]&ds=[Desc]&dt=[Dt]
```

## Categories

**Income** — Salary, E-transfer, Refund, Other

**Expense** — Food, Transport, Bills, Shopping, Rent, Going-out,
Entertainment, Groceries, Other

Menu labels can carry spaces and emoji. The Text action's value cannot.

## Adding a category

1. Add the item to the relevant Choose from Menu
2. Inside its branch, add a **Text** action with the value
3. Add **Set Variable** `Cat` pointing at that Text

Nothing changes on the script side — categories are read from the data.

## Debugging on device

Add a **Quick Look** after Get Contents of URL to see the response. Point it at
the **Text** action instead to see the URL being sent.

Add `echo=1&` immediately after `/exec?` to have the server report what it
received without writing anything.
