# JCBudgeting Standalone Server

This server can run without the desktop app on another computer.

## Packaging Targets

- Windows publish: `win-x64`
- Linux publish: `linux-x64`
- Linux ARM publish: `linux-arm64`
- Container support: `Dockerfile` included in the server project folder for Linux/container hosts such as Umbrel-style setups

Linux publishes are now built as self-contained packages, so Ubuntu does not need a separate .NET runtime installed just to run the server.

## Quick Start

1. Copy `budgetserver.local.json.example` to `budgetserver.local.json`.
2. Edit `BudgetServer.DatabasePath` to point at the `.jcbdb` file on that machine.
3. Choose one of these two configuration styles:
   - Set `BudgetServer.SettingsPath` to a real `usersettings.ini` file copied from the desktop app.
   - Or leave `SettingsPath` blank and fill in:
     - `BudgetPeriod`
     - `BudgetStartDate`
     - `BudgetYears`
4. Start the server:
   - `JCBudgeting.Server.exe`
5. Open the health endpoint in a browser:
   - `http://<server-host>:5099/api/health`

## Linux Quick Start

1. Build a Linux package:
   - `powershell -ExecutionPolicy Bypass -File .\scripts\Build-StandaloneServer.ps1 -RuntimeIds linux-x64`
   - or:
   - `powershell -ExecutionPolicy Bypass -File .\scripts\Build-StandaloneServer.ps1 -RuntimeIds linux-arm64`
2. Copy the output folder for that runtime to the Linux machine.
3. Make sure the launcher script is executable:
   - `chmod +x ./start-server.sh`
4. Start it:
   - `./start-server.sh`
5. Browse to:
   - `http://<server-host>:5099/setup`

If you prefer to launch the binary directly:
   - `chmod +x ./JCBudgeting.Server`
   - `./JCBudgeting.Server`

For easier first-run testing on Ubuntu, you can also use:
- `./start-server.sh`

That script will set the execute bit on `JCBudgeting.Server` and then launch it.

## Docker Quick Start

1. Build a Linux publish first, for example `linux-x64`.
2. Copy the generated publish contents plus the included `Dockerfile` into your container build context.
3. Build the image:
   - `docker build -t jcbudgeting-server .`
4. Run it:
   - `docker run -p 5099:5099 -v /your/data/path:/app/Databases jcbudgeting-server`

For container use, keeping the `Databases` folder on a mounted volume is the safer long-term path so server updates do not replace user data.

## Notes

- `appsettings.json` contains the base defaults.
- `budgetserver.local.json` is the machine-specific override file.
- You can also override values with environment variables:
  - `BudgetServer__DatabasePath`
  - `BudgetServer__SettingsPath`
  - `BudgetServer__BudgetPeriod`
  - `BudgetServer__BudgetStartDate`
  - `BudgetServer__BudgetYears`
  - `ASPNETCORE_URLS`

## Typical Standalone Config

```json
{
  "Urls": "http://0.0.0.0:5099",
  "BudgetServer": {
    "DatabasePath": "D:\\BudgetData\\FamilyBudget.jcbdb",
    "SettingsPath": "",
    "BudgetPeriod": "Monthly",
    "BudgetStartDate": "04/09/2026",
    "BudgetYears": 20
  }
}
```

## Desktop-managed Mode

The desktop app can still start this same server and inject the database/settings paths automatically. The standalone configuration files do not break that workflow.
