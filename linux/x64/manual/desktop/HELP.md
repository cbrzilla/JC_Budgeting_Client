# JCBudgeting User Guide

## What JCBudgeting Does
JCBudgeting is a budgeting program built around a forward-looking budget sheet. It helps you:

- manage accounts, income, expenses, savings goals, debts, and transactions
- plan future periods instead of only looking backward
- see budget health in charts and flow views
- work locally or connect to a JC Server
- keep working from an offline cache if the server becomes unavailable

Jumper Cable (JC) Budgeting helps you track, plan, and forecast your budget across multiple accounts, incomes, bills, savings, debts, and expenses together in one budget workspace.

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
- `Data`
- `Settings`

Each area is described below.

## Getting Started

### Local Database Workflow
Use this when you want everything on one machine.

1. Open `JCBudgeting`.
2. Go to `Settings`.
3. Set `Database Source` to `Local Only (No External Access)`.
4. Under `Current Database`, use `Load` to open an existing `.jcbdb` file or `Create New` to make a new one.
5. Set your budget timeline:
   - `Budget Start`
   - `Budget Period`
   - `Years`
6. Start adding accounts, income, savings, debts, expenses, and transactions.

### Server Workflow
Use this when the database lives on a JC Server and the client connects to it.

1. Start the JC Server.
2. Use the server page at `http://<server-host>:5099/server`.
3. On the server page, create or select the database the server should use.
4. Open the desktop client.
5. Go to `Settings`.
6. Set `Database Source` to `External Host`.
7. Enter the server host name or IP address and port.
8. Click `Connect` if needed.
9. Use `Open Server Page` anytime you want to reopen the server page.

Important:

- The server page is for choosing or creating the server database, reviewing server settings, viewing server logs, and checking standalone server updates.
- `Upload DB` and `Download DB` are done from the desktop client in `Settings`.

### Hosted Server On This Computer
This mode is useful when you want the desktop app and a server running on the same machine.

- Choose `Local (External Access Allowed)` in `Settings`.
- The desktop app hosts the currently loaded local database so phones or other clients can connect to this computer.
- If you switch back to `Local Only (No External Access)`, the hosted background server should stop.
- Other clients can then connect to this computer.

## First-Time Guidance
On first use, JCBudgeting can show a guided tour to help explain the main tabs and setup flow.

- The startup tour begins in `Settings` so you can create, load, or connect to a database first.
- Several editor tabs also have their own detailed tours the first time you create your first item there.
- You can reopen tours later with the `Show Guided Tour` or `Show Tour` buttons where available.

## Understanding Server Status
When you are using a server, the app shows a status pill in the footer and related status text in the budget area.

Possible states:

- `Online`
- `Offline`
- `Online, but no database loaded`

If the server goes offline during a session, the app can offer:

- `Retry Server`
- `Switch To Offline`
- `Close Software`

If you switch to offline mode and later reconnect to the server, use `Upload DB` to send your offline changes back to the server database.

## Versions And Updates
The desktop client checks the JC Server version before connecting. If the server requires a newer or incompatible client, the client blocks the connection before reading or writing budget data.

The desktop app can check the client download repo for updates from the `Settings` help area. Standalone servers that are not managed by Umbrel can check the same download repo from the server page at `http://<server-host>:5099/server`.

## Overview Tab
The `Overview` tab gives you a high-level picture of the current budget period and future trends.

It includes these sub-sections:

- `Summary`
- `Budget Distribution`
- `Savings`
- `Comparison`

### Summary
Shows major budget metrics such as:

- income
- planned outflow
- savings contributions
- total savings
- new debt charges
- net flow
- current debt balance

It also includes charts for:

- projected account balances
- savings balances
- projected debt balances

### Budget Distribution
Shows a Sankey-style flow diagram of how budget money moves through the budget.

Use this to understand:

- where money is coming from
- where money is being assigned
- how large each flow is relative to others

### Savings
Shows a savings flow diagram so you can see:

- starting balances
- contributions
- expenses funded from savings
- remaining savings balances

### Comparison
Compares the selected period to prior matching periods so you can quickly see whether planned amounts are above, below, or near historical averages.

## Budget Tab
The `Budget` tab is the core planning sheet.

It lets you:

- view the budget across multiple periods
- scroll through future periods
- see grouped rows for accounts, income, expenses, savings, and debts
- use sticky account rows for easier reading
- edit a budget cell directly from the sheet
- export the budget

At the top of the budget area, a color key explains the budget text colors:

- `Blue`: has transactions linked
- `Purple`: mixture of linked transactions and manual adjustments
- `Red`: manually overriden
- `Black` or `White` (depending on theme): calculated

### Budget Cell Editing
When you click a budget cell, the editor lets you adjust how that value is controlled.

Depending on the item, you can:

- manually set a value
- use a calculated value
- add or subtract an adjustment
- mark items paid when applicable
- save notes
- clear the override and return to the normal behavior

Explicit zeros are preserved when you intentionally set them, including calculated zero values.

### Hidden Items
Hidden items are still tracked clearly in the budget:

- hidden rows are labeled with `(Hidden)` when relevant
- group labels show hidden counts
- parent sections stay visible even if only hidden children remain

## Accounts Tab
Use `Accounts` to manage the funding sources that power the budget.

Typical fields include:

- account name
- account type
- category
- starting balance
- notes

Accounts are used throughout the rest of the app for budgeting and funding references.

Until at least one account exists, the dependent tabs stay blocked:

- `Income`
- `Savings`
- `Debts`
- `Expenses`
- `Transactions`

## Income Tab
Use `Income` to define recurring or manual income sources.

Income supports patterns such as:

- weekly
- bi-weekly
- monthly on day
- yearly on date
- same as another item
- manually entered

You can also control when each item starts and stops affecting the budget.

Income stays unavailable until at least one account exists.

## Savings Tab
Use `Savings` for savings goals, sinking funds, emergency funds, allowances, and similar buckets.

Savings items can include:

- description
- category
- funding account
- deposit schedule
- goal amount
- target date
- hidden or active status
- notes

Savings appear in both the budget sheet and the overview savings flow.

Savings stays unavailable until at least one account exists.

## Debts Tab
Use `Debts` to manage loans, cards, and other liabilities.

Debt records can include:

- description
- lender
- debt type
- APR
- current balance
- original principal
- minimum payment
- due information
- funding account
- notes

The app uses debt information to project balances and show debt trends in `Overview`.

Debts stays unavailable until at least one account exists.

## Expenses Tab
Use `Expenses` to manage recurring and manual bills or spending categories.

Expenses support:

- description
- category
- amount
- cadence
- due day or date
- start and end timing
- funding source
- same-as relationships
- hidden or active status
- notes

Expenses can be funded from:

- accounts
- savings
- debts

Those funding choices affect the budget flow and savings flow diagrams.

Expenses stays unavailable until at least one account exists.

## Transactions Tab
Use `Transactions` to import and assign actual transaction activity.

Key features:

- import transactions
- view source, date, description, amount, notes
- assign a transaction to one or more budget items
- use quick assign where supported
- see unlinked or split assignments clearly

Import note:

- Because each financial institution can export transactions differently, JC Budgeting expects negative amounts as debits and positive amounts as credits.
- You may need to adjust the import file before importing so it matches that format.

This area is especially useful for tying real-world activity back to the budget plan.

Transactions stays unavailable until at least one account exists.

## Data Tab
The `Data` tab is a raw read-only database inspection area.

Use it to:

- view database tables directly
- inspect saved rows and values
- verify what was actually written to the database

This is mainly for troubleshooting and advanced checking.

## Settings Tab
The `Settings` tab controls how the app connects, where data comes from, and how the budget timeline is defined.

Main settings include:

- database source
- external host address
- external host port
- open server setup
- upload DB
- download DB
- current database message
- local load or backup actions
- budget start date
- budget period
- number of years
- theme
- theme color

`Database Source` has three modes:

- `Local Only (No External Access)`: keeps the database on this computer only
- `Local (External Access Allowed)`: keeps the database on this computer and hosts it for other devices
- `External Host`: connects to an existing JC Server

If no active database is loaded, `Settings` stays usable and the other tabs are blocked until you finish setup.

## Uploading And Downloading Databases
When using a server connection:

- `Upload DB` sends a local `.jcbdb` file from the client to the server
- `Download DB` pulls the currently active server database down to the client machine

This is especially useful when:

- moving work between local and server environments
- restoring a copy
- uploading offline work back to the server later

## Offline Mode
If the server becomes unavailable and you have a cached copy, JCBudgeting can switch to offline mode.

Important behavior:

- switching to offline mode alone does not mark the cache as needing upload
- the cache is marked for upload only after actual offline changes are made
- when the server becomes available again, the app can prompt you to upload offline changes
- the unavailable popup includes a `Close Software` option

Recommended workflow after offline work:

1. Reconnect to the server.
2. Go to `Settings`.
3. Use `Upload DB` to send the updated offline database back to the server.

## Logs
Both the desktop app and the server can write logs for troubleshooting.

Look for a `Logs` folder near the running application.

Typical uses:

- checking startup failures
- checking server connection issues
- reviewing runtime errors

## Tips For New Users

- Start in `Settings` and get the database source and active database working before trying to enter other data.
- Set up `Accounts`, `Income`, `Savings`, `Debts`, and `Expenses` before doing deep budget edits.
- Use `Overview` to validate whether the budget still makes sense after major changes.
- Use the `Transactions` tab regularly so actual activity stays tied to the plan.
- If you are sharing a budget on another machine, use `External Host` mode instead of pointing multiple people at the same file directly.
- If you run the Linux server from a shared or mounted folder, keep the live database itself in a normal local Linux folder to avoid SQLite write issues.

## Troubleshooting

### The client says the server is online but no database is loaded
Open the server setup page and create or select a database for the server.

### The other tabs are blocked

- If the message says `No Active Database`, go to `Settings` and create, load, or connect to a database.
- If the message says `No Accounts Created Yet`, go to `Accounts` and create an account first.

### The server cannot be reached

- verify the server is running
- verify the host/IP and port in `Settings`
- make sure the machine firewall allows the chosen port
- use `Retry Server` or switch to offline mode if needed

### My offline changes are not on the server yet
Reconnect to the server and use `Upload DB` from the client.

### I want to inspect exactly what is in the database
Use the `Data` tab and load the database tables.

## Quick Start Checklist

1. Create or connect to a database.
2. Confirm the budget timeline in `Settings`.
3. Add accounts.
4. Add income.
5. Add savings items.
6. Add debts.
7. Add expenses.
8. Review the `Budget` tab.
9. Review the `Overview` tab.
10. Import and assign transactions as needed.
