# Setup

Two halves. Do the sheet first — you need its URL before the Shortcut is useful.

---

## Part 1 — Google Sheet

### 1. Create the spreadsheet

Go to [sheets.google.com](https://sheets.google.com), create a blank spreadsheet, name it whatever you like.

### 2. Add the script

**Extensions → Apps Script.** Delete whatever is in the editor, paste the entire contents of [`src/Code.gs`](../src/Code.gs), then save.

### 3. Deploy as a web app

1. **Deploy → New deployment**
2. Click the gear next to "Select type" → **Web app**
3. Set **Execute as: Me** and **Who has access: Anyone**
4. **Deploy**
5. Authorise it. Google will warn that the app is unverified — it is your own script, so click **Advanced → Go to project**.
6. Copy the **Web app URL**. It ends in `/exec`.

> **Keep this URL private.** Anyone holding it can write rows into your sheet. There is no login on it.

### 4. Check it responds

Paste the URL into a browser. You should see:

```
v1.0 — Expense Tracker API is running. No amount parameter was received.
```

Then try a real entry:

```
YOUR_URL?am=12&ty=Expense&ct=Food&ds=test&dt=2026-08-15
```

A success JSON comes back and an `August 2026` tab appears. Delete that test row afterwards using **💰 Tracker → Delete selected entries**.

### Updating the script later

Editing and saving does **not** change what the live URL serves. Every time:

**Deploy → Manage deployments → ✏️ pencil → Version: New version → Deploy**

Use "New deployment" only if you want a second, separate URL.

---

## Part 2 — iOS Shortcut

Open **Shortcuts**, tap **+**, name it `Add Expense`.

### Structure

Indentation shows nesting. Build it in this order.

```
Ask for Number — "Enter Amount"
Set variable Amt to Ask for Input

Choose from Menu — "Income or Expense?"

    Income
        Text: Income
        Set variable Typ to Text
        Choose from Menu — "Category"
            Salary
                Text: Salary
                Set variable Cat to Text
            E-transfer
                Text: E-transfer
                Set variable Cat to Text
            Refund
                Text: Refund
                Set variable Cat to Text
            Other
                Text: Other
                Set variable Cat to Text
        End Menu

    Expense
        Text: Expense
        Set variable Typ to Text
        Choose from Menu — "Category"
            Food / Transport / Bills / Shopping / Rent /
            Going-out / Entertainment / Groceries / Other
              (each branch: Text with the name, then Set variable Cat to Text)
        End Menu

End Menu

Ask for Input (Date) — "Date of expense?"
Format Date — Custom: yyyy-MM-dd
Set variable Dt to Formatted Date

Ask for Text — "Description"
URL Encode — Ask for Input
Set variable Desc to URL Encoded Text

Text — [the request URL, see below]
Get Contents of URL — [Text]
Show Notification
```

### The request URL

Build this inside a **Text** action, not directly in Get Contents of URL. Type the plain text and insert variables where shown in brackets.

```
https://script.google.com/macros/s/YOUR_ID/exec?am=[Amt]&ty=[Typ]&ct=[Cat]&ds=[Desc]&dt=[Dt]
```

Then set **Get Contents of URL** to that Text variable, method GET, headers empty.

### Points that catch people out

**Set Variable needs a Text action.** Its input field only accepts variables, so you cannot type a literal value into it. Put the word in a Text action first, then point Set Variable at that.

**Menu Result is ambiguous.** With more than one menu, a generic `Menu Result` pill may bind to the wrong one. Naming each value through Text plus Set Variable, inside its own branch, removes the guesswork.

**Keep category values single-word.** The menu label can say `Going out 🆒` with a space and an emoji; the Text action's value must be `Going-out`. Category values are not URL-encoded, so a space truncates the request.

**Capital `MM` in the date format.** Lowercase `mm` is minutes in Shortcuts, so `yyyy-mm-dd` gives you the wrong value.

**Build the URL in a Text action.** A long URL typed directly into Get Contents of URL can be truncated in transit, dropping whatever comes last.

### Bind it to Back Tap

**Settings → Accessibility → Touch → Back Tap → Double Tap** → scroll to the bottom → select **Add Expense**.

Back Tap can be finicky. Tap a bit more firmly than feels natural, around the middle of the back of the phone, and try a slightly different spot if nothing happens. Thick cases interfere.

---

## Part 3 — Optional extras

### The Tracker menu

Reopen the spreadsheet and a **💰 Tracker** menu appears next to Help:

- **Refresh dashboard** — rebuild everything, faster than calling `?repair=1`
- **Delete selected entries** — select rows on a month tab, then use this instead of deleting rows manually
- **Enable auto-refresh** — installs a trigger that rebuilds after any manual edit

The first use of these prompts for additional permissions, because menus and triggers need broader scopes than the web app deployment does.

### Time zone

Month tabs are named from the **spreadsheet's** time zone, not your phone's. If an entry logged near midnight on the 1st lands in the wrong month, set **File → Settings → Time zone** to match where you are.
