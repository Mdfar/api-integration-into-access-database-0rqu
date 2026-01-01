Apify to MS Access Connector
Setup

Open MS Access and press Alt+F11.

Go to Tools > References and check Microsoft WinHTTP Services, v5.1.

Import Apify_Integration.bas and Credit_Tracker.bas.

Update APIFY_TOKEN with your credentials.

Features

Asynchronous Execution: Triggers the actor and waits for completion.

Credit Safeguard: Logs credit usage before and after every run.

Relational Integrity: Links every lead to a specific ApifyRunID for audit trails.