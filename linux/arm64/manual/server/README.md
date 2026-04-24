# JCBudgeting Server (linux arm64)

## How To Run
1. On Ubuntu/Linux, either right-click 'start-server.sh' and choose 'Run as a Program', or run './start-server.sh' after 'chmod +x ./start-server.sh' if needed.
2. Wait for the server window to show the sharing host/IP, port, setup page link, and active database name.

## Server Setup Page
Once the server is running, open this in a browser on the server machine or another device on the same network:

- http://<server-host>:5099/setup

Replace <server-host> with the host name or IP shown in the server window.

Use the setup page to:
- choose or create the database
- confirm the server port and budget timeline settings
- review the current host/IP and setup details shown by the server

Database upload and download are done from the JCBudgeting client, not from the server setup page.

## Notes
- The server window also shows licensing and project information at startup.
- Install JC Budgeting clients from: https://github.com/cbrzilla/JC_Budgeting_Client
