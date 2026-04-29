# JCBudgeting User Guide

## What JC Budgeting Does
Jumper Cable (JC) Budgeting helps you track, plan, and forecast your budget across multiple accounts, incomes, bills, savings, debts, expenses, and transactions together in one budget workspace.

JC Budgeting is built around a forward-looking budget sheet. Instead of only recording the past, it helps you plan future periods and understand how your choices affect balances, savings, debt, and overall sustainability.

Main capabilities include:

- managing accounts, income, savings, debts, expenses, and transactions
- planning future budget periods in a spreadsheet-style budget sheet
- comparing calculated amounts, linked transaction amounts, and manual overrides
- viewing account, savings, debt, and distribution visuals in `Overview`
- working locally, through a local server, or through an external JC Server
- continuing from an offline cache if a server becomes temporarily unavailable

  [![Watch the video]](https://youtu.be/zbjzAUdrEmU)


## Main Areas Of The Program
The desktop app is organized into these main tabs:

- `Overview`
- `Budget`
- `Accounts`
- `Income`
- `Savings`
- `Debts`
- `Expenses`
- `Transactions`
- `Settings`

## Getting Started

### First-Time Launch
On first launch, JC Budgeting can show a guided welcome tour.

The startup tour:

- begins in `Settings`
- explains the database source choices
- explains how to create, load, or connect to a database
- explains theme customization and the main workflow

If you skip or complete the tour, it should not automatically reappear. You can still reopen it later from the `Help` area in `Settings`.

### Important Setup Rule
You must have an active database before the main planning tabs can be used.

If there is no active database:

- `Settings` stays usable
- the other tabs are blocked with a `No Active Database` overlay
- the budget status banner shows `No Active Database`

You also need at least one account before certain tabs can be used normally. If no accounts exist yet, dependent tabs show a `No Accounts Created Yet` overlay.

## Database Source Options
The `Settings` tab uses `Database Source` to choose where your budget data comes from.

There are three modes:

- `Local Only`: stores the database on this computer only
- `Local Server`: uses a JC Server running on this same computer so other devices can connect to it
- `External Server`: connects to a JC Server running somewhere else

### Local Only
Use this when you want the entire budget to stay on one computer with no server access from other devices.

Typical workflow:

1. Open `Settings`.
2. Set `Database Source` to `Local Only`.
3. Use `Load` to open an existing `.jcbdb` file or `Create New` to make a new one.
4. Set the budget timeline:
   - `Budget Start Date`
   - `Budget Period`
   - `Budget Years`
5. Start adding your budgeting data.

### Local Server
Use this when you want the client and the server to run on the same machine.

Typical behavior:

- `Server Address` is usually `localhost`
- the client can try to start the local server automatically
- if the local server is not running, the app can show `Start Server`
- if the server is not installed yet, the app can show `Install Server`
- if you already had a local database active, the app can transfer it into the local server flow when needed

In this mode, the server controls the shared budget data similarly to an external server.

### External Server
Use this when your data is hosted by a JC Server on another computer or device.

Typical workflow:

1. Start the standalone JC Server.
2. Open the server web page at `http://<server-host>:5099/server`.
3. Create or select the active database on the server (or upload an existing database to the server using the client).
4. Open the desktop client.
5. Set `Database Source` to `External Server`.
6. Enter `Server Address` and `Port`.
7. Connect to the server.

Important:

- the server web page is used for server setup, live status, logs, and update information
- database upload and download are done from the desktop client, not the server web page

## Guided Tours And Help
JC Budgeting includes several guided tours.

### Startup Tour
The startup tour gives a high-level walkthrough of:

- what the app does
- where to begin
- database source choices
- local versus server setup
- theme customization
- the main tab order

### Contextual Editor Tours
Detailed editor tours can appear the first time you create your first item in certain tabs:

- `Accounts`
- `Income`
- `Savings`
- `Debts`
- `Expenses`
- `Transactions`

These tours explain the key inputs in the editor and how those values affect the budget.

You can reopen them later with `Show Tour` where available. The button is disabled for tabs that need an item selected first.

### Help Area
The `Settings` tab includes a `Help` area where you can:

- reopen the startup guided tour
- check for client updates
- open the client repository link for more information
- open the support link

## Server Status And Compatibility
When a server mode is active, JC Budgeting shows a server status indicator in the footer and related status text in the budget area.

Possible states include:

- `JC Server Online`
- `JC Server Offline`
- `JC Server No Database`
- `JC Server Version Mismatch`

If the server version is incompatible, the desktop client blocks the connection before reading or writing budget data.

The compatibility check compares:

- client version
- client compatibility series
- server version
- server minimum supported client version

## Updates
JC Budgeting can check GitHub-hosted releases for updates.

### Client Update Check
The desktop client can check the client download repo from the `Help` area in `Settings`.

### Server Update Check
Standalone servers can check for updates from the server web page.

Umbrel-managed servers use the Umbrel update flow instead of the standalone server flow.

## Offline Mode
If you are using a server and the server becomes unavailable, JC Budgeting can use an offline cache.

If the server drops during use, the client can offer:

- `Retry Server`
- `Switch To Offline`
- `Close Software`

Important offline behavior:

- switching to offline mode by itself does not mark the cache as needing upload
- the cache is marked for upload only after actual offline edits are made
- when the server becomes available again, the client can prompt you to upload local offline changes

Recommended workflow after offline work:

1. Reconnect to the server.
2. Open `Settings`.
3. Use `Upload DB` to send the updated offline database back to the server.

## Overview Tab
The `Overview` tab gives a high-level picture of the current and future budget.

Main sections include:

- `Summary`
- `Budget Distribution`
- `Savings`
- `Comparison`

### Summary
Shows major measures such as:

- income
- planned outflow
- savings contributions
- total savings
- new debt charges
- net flow
- current debt balance

It also includes projected charts for:

- account balances
- savings balances
- debt balances

### Budget Distribution
Shows a Sankey-style distribution view of how money moves through the budget.

Use it to see:

- where money is sourced
- where money is assigned
- relative flow size between categories

### Savings
Shows a savings Sankey view so you can understand:

- starting balances
- planned contributions
- spending funded from savings
- projected remaining savings balances

### Comparison
Compares the selected period to prior matching periods so you can see whether current planned amounts are above, below, or near historical averages.

## Budget Tab
The `Budget` tab is the core planning sheet.

It lets you:

- view multiple periods at once
- scroll through current and future periods
- work with grouped rows for accounts, income, savings, debts, and expenses
- use hidden rows without losing visibility of hidden counts
- open a budget cell editor directly from the sheet
- export the budget

### Budget Text Color Key
The top budget area includes a color key for text in the budget sheet:

- `Blue`: Has Transactions Linked
- `Purple`: Mixture of Linked Transactions & Manual Adjustments
- `Red`: Manually Overriden
- `Black` or `White` depending on theme: Calculated

### Budget Cell Editing
When you click a budget cell, the editor lets you control how the value behaves.

Depending on the row type, you can:

- manually set a value
- use a calculated value
- add an adjustment
- mark an item paid when applicable
- save notes
- reset the override

Explicit zero values are preserved when you intentionally set them, including calculated zero values.

### Hidden Rows
Hidden budget items are still represented clearly:

- hidden rows are labeled with `(Hidden)` where relevant
- parent labels show hidden counts
- parent sections remain visible even if only hidden children remain

## Accounts Tab
Use `Accounts` to define the funding sources used throughout the budget.

Common fields include:

- name
- account type
- safety net
- hidden
- active
- login link
- notes

Accounts are the foundation for the rest of the planning flow. Until at least one account exists, several other tabs stay blocked.

## Income Tab
Use `Income` to define recurring or manual income sources.

Income supports timing patterns such as:

- weekly
- bi-weekly
- monthly
- yearly
- same-as patterns where supported
- manual entry

Income also supports timing controls such as start and end behavior so income can begin or stop affecting future periods.

## Savings Tab
Use `Savings` for savings goals, sinking funds, reserve buckets, and similar planned balances.

Savings records can include:

- description
- category
- funding source
- contribution amount
- frequency
- start date
- goal amount
- goal date
- hidden
- active
- login link
- notes

Savings feed both the budget sheet and the overview savings visual.

## Debts Tab
Use `Debts` to manage liabilities such as loans, cards, and other balances owed.

Debt records can include:

- description
- lender or institution
- debt type
- loan type
- APR
- current balance
- original principal
- payment amount
- frequency
- due timing
- funding account
- hidden
- active
- login link
- notes

Debt settings affect balance projections and debt visuals in `Overview`.

## Expenses Tab
Use `Expenses` to manage recurring bills and flexible spending categories.

Expense records can include:

- description
- category
- amount
- cadence or frequency
- due timing
- start or stop timing
- funding source
- same-as relationships where applicable
- hidden
- active
- login link
- notes

Categories can be selected from existing choices, and you can also type your own category text when needed.

Expenses can be funded from:

- accounts
- savings
- debts

Those choices affect the budget and overview distribution flows.

## Transactions Tab
Use `Transactions` to import and assign real-world activity.

Main features include:

- importing transaction files
- reviewing source, date, description, amount, and notes
- assigning transactions to one or more budget items
- using quick assign where supported
- tracking split assignments clearly

Important import note:

- JC Budgeting expects negative amounts as debits and positive amounts as credits
- some financial institutions export data differently, so you may need to edit the file before import so it matches that format

## Settings Tab
The `Settings` tab controls connection mode, database actions, timeline settings, theme choices, and help.

Main areas include:

- `Database Source`
- server connection details
- current database status message
- load, create, backup, upload, and download actions
- budget timeline settings
- theme and accent color settings
- help and update-check options

### Database Actions
Depending on the selected mode, `Settings` can provide:

- `Load`
- `Create New`
- `Backup`
- `Connect`
- `Open Server Page`
- `Upload DB`
- `Download DB`
- `Install Server`
- `Start Server`
- `Shut Down Server`

The current database display is a status/message field rather than an editable text field.

### Timeline Settings
The budget timeline is defined by:

- `Budget Period`
- `Budget Start Date`
- `Budget Years`

When you are connected through a server mode, these settings are controlled by the active server database.

### Theme Settings
The `Settings` tab also includes theme customization options such as:

- light or dark theme
- accent or theme color choices
- animation behavior

## Server Web Page
The standalone JC Server includes a browser-based server page at:

`http://<server-host>:5099/server`

The server page can show:

- current active database
- primary access address
- setup/server page address
- client download link
- live server log output
- server version and update information
- copyright and usage notices

Use the server page to:

- create or select the server database
- review server settings
- confirm the address clients should use
- monitor runtime activity and logs

## Logging
Both the desktop app and the server can write logs for troubleshooting.

Look for a `Logs` folder near the running application or installed app data location.

Typical uses:

- checking startup failures
- checking connection or compatibility issues
- reviewing runtime errors
- debugging deployment-specific issues

## Tips For New Users

- Start in `Settings` and get the active database working before entering data.
- Create accounts first, because several other tabs depend on them.
- Use the guided tours when they appear, especially on first setup and first item creation.
- Use `Overview` to confirm your plan still makes sense after major edits.
- Use `Transactions` regularly so actual activity stays tied to the planned budget.
- If several devices need to share the same budget, use `Local Server` or `External Server` instead of pointing multiple clients directly at one file.
- On Linux, keep live SQLite databases in a normal local filesystem path instead of certain shared or mounted VM folders to avoid SQLite read-only or disk I/O problems.

## Troubleshooting

### The client says no active database is loaded
Go to `Settings` and create, load, or connect to a database.

### The other tabs are blocked

- If the overlay says `No Active Database`, finish database setup in `Settings`
- If the overlay says `No Accounts Created Yet`, create an account in `Accounts` first

### The server says it is online but no database is loaded
Open the server page and create or select a database for the server.

### The server cannot be reached

- verify the server is running
- verify `Server Address` and `Port`
- make sure the machine firewall allows the selected port
- use `Retry Server`, `Switch To Offline`, or `Close Software` as needed

### I made offline changes and the server does not have them yet
Reconnect to the server and use `Upload DB` from the client.

### I want to inspect the raw saved data
Use logs, database backups, or the server/client troubleshooting tools to inspect what was saved.

### The client says there is a server version mismatch
Update the client and/or server so the client meets the server's minimum compatible version requirements.

## Quick Start Checklist

1. Open `Settings`.
2. Choose the correct `Database Source`.
3. Create, load, or connect to a database.
4. Confirm the budget timeline settings.
5. Add accounts.
6. Add income.
7. Add savings items.
8. Add debts.
9. Add expenses.
10. Review the `Budget` tab.
11. Review the `Overview` tab.
12. Import and assign transactions as needed.
