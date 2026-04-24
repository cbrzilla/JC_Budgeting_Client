# Linux Installers

Install the client:

    sudo apt install ./jcbudgeting-client_1.0.0_amd64.deb

Launch the client:

    jcbudgeting
    
Or search for:

    JC Budgeting

Uninstall the client:

    sudo apt remove jcbudgeting-client

Install the server:

    sudo apt install ./jcbudgeting-server_1.0.0_amd64.deb

The server package installs and starts the jcbudgeting-server systemd service. Open http://localhost:5099/server after install.

Uninstall the server:

    sudo apt remove jcbudgeting-server

The uninstall removes the app/service package, but it does not intentionally delete user-created databases.