let smarClient;           // Smartsheet JS client object

// Dependent libraries
const express = require("express");
const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb' }));
const bodyParser = require("body-parser");
app.use(bodyParser.json());

const smarSdk = require("smartsheet");

// Initialize client SDK
function initializeSmartsheetClient(token, logLevel) {
    smarClient = smarSdk.createClient({
        // If token is falsy, value will be read from SMARTSHEET_ACCESS_TOKEN environment variable
        accessToken: token,
        logLevel: logLevel
    });
}

// Check that we can access the sheet
async function probeSheet(targetSheetId) {
    console.log(`Checking for sheet id: ${targetSheetId}`);
    const getSheetOptions = {
        id: targetSheetId,
        queryParameters: { pageSize: 1 } // Only return first row to reduce payload
    };
    const sheetResponse = await smarClient.sheets.getSheet(getSheetOptions);
    console.log(`Found sheet: "${sheetResponse.name}" at ${sheetResponse.permalink}`);
}

/*
* A webhook only needs to be created once.
* But hooks will be disabled if validation or callbacks fail.
* This method looks for an existing matching hook to reuse, else creates a new one.
*/
async function initializeHook(targetSheetId, hookName, callbackUrl) {
    try {
        let webhook = null;

        // Get *all* my hooks
        const listHooksResponse = await smarClient.webhooks.listWebhooks({
            includeAll: true
        });
        console.log(`Found ${listHooksResponse.totalCount} hooks owned by user`);

        // Check for existing hooks on this sheet for this app
        for (const hook of listHooksResponse.data) {
            if (hook.scopeObjectId === targetSheetId
                && hook.name === hookName
                // && hook.callbackUrl === callbackUrl   
            ) {
                webhook = hook;
                console.log(`Found matching hook with id: ${webhook.id}`);
                break;
            }
        }

        if (!webhook) {
            // Can't use any existing hook - create a new one
            const options = {
                body: {
                    name: hookName,
                    callbackUrl,
                    scope: "sheet",
                    scopeObjectId: targetSheetId,
                    events: ["*.*"],
                    version: 1
                }
            };

            const createResponse = await smarClient.webhooks.createWebhook(options);
            webhook = createResponse.result;

            console.log(`Created new hook: ${webhook.id}`);
        }

        // Make sure webhook is enabled and pointing to our current url
        const options = {
            webhookId: webhook.id,
            callbackUrl: callbackUrl,
            body: { enabled: true }
        };

        const updateResponse = await smarClient.webhooks.updateWebhook(options);
        const updatedWebhook = updateResponse.result;
        
        console.log(`Hook enabled: ${updatedWebhook.enabled}, status: ${updatedWebhook.status}`);
    } catch (err) {
        console.error(err);
    }
}


// This method receives the webhook callbacks from Smartsheet
app.post("/", async (req, res) => {
    try {
        const body = req.body;

        // Callback could be due to validation, status change, or actual sheet change events
        if (body.challenge) {
            console.log("Received verification callback");
            // Verify we are listening by echoing challenge value
            res.status(200)
                .json({ smartsheetHookResponse: body.challenge });
        } else if (body.events) {
            console.log(`Received event callback with ${body.events.length} events at ${new Date().toLocaleString()}`);
            processedEventIds.clear();
            await processEvents(body);
            res.sendStatus(200);
        }
        else if (body.newWebHookStatus) {
            console.log(`Received status callback, new status: ${body.newWebHookStatus}`);
            res.sendStatus(200);
        } else {
            console.log(`Received unknown callback: ${body}`);
            res.sendStatus(200);
        }
    } catch (error) {
        console.log(error);
        res.status(500).send(`Error: ${error}`);
    }
});


// let arg1, arg2, arg3;

// async function run(arg1, arg2, arg3) {
//     const { spawn } = require('child_process');
//     const pythonProcess = spawn('python', ['new.py', arg1, arg2, arg3]);
//     pythonProcess.stderr.on('data', (error) => {
//         console.error('Error fetching data:', error.toString());
//     });
// }

let processedEventIds = new Set();

async function processEvents(callbackData) {
    if (callbackData.scope !== "sheet") {
        return;
    }

    for (const event of callbackData.events) {
        // Add event ID to set

        if (event.objectType == "row") {
            console.log(`Row: ${event.eventType}, row id: ${event.id}`);

            try {
                if (event.eventType == "deleted") {
                    console.error('Row deleted in SS');
                    arg1 = event.eventType;
                    arg2 = event.id;
                    arg3 = "output";
                    run(arg1, arg2, arg3)
                } else if (event.eventType == "updated") {
                    const eventId = `${event.id}_${event.version}`; // Combine ID and version
                    if (processedEventIds.has(eventId)) {
                        // console.log(`Skipping duplicate event: ${eventId}`);
                        continue; // Skip processing duplicate events
                    }

                    processedEventIds.add(eventId);
                    console.log('Row updated in SS');
                    // Fetch the updated row and process the changes here
                    const updatedRow = await smarClient.sheets.getRow({
                        sheetId: 8199960751198084, // Assuming event.sheetId contains the sheet ID
                        rowId: event.id // Use the row ID from the webhook event
                    });
                    const values = [];
                    for (const a of updatedRow.cells) {
                        const value = "'" + a.value + "'";
                        values.push(value);
                    }
                    values.push(event.id)
                    const output = values.join(', ');
                    arg1 = event.eventType;
                    arg2 = event.id;
                    arg3 = output;
                    await run(arg1, arg2, arg3)
                    // Process the updated row as needed
                } else {
                    const eventId = `${event.id}_${event.version}`; // Combine ID and version
                    if (processedEventIds.has(eventId)) {
                        // console.log(`Skipping duplicate event: ${eventId}`);
                        continue; // Skip processing duplicate events
                    }

                    processedEventIds.add(eventId);
                    const row = await smarClient.sheets.getRow({
                        sheetId: 8199960751198084, // Assuming event.sheetId contains the sheet ID
                        rowId: event.id // Use the row ID from the webhook event
                    });

                    const values = [];
                    for (const a of row.cells) {
                        const value = "'" + a.value + "'";
                        values.push(value);
                    }
                    values.push(event.id)
                    const output = values.join(', ');
                    arg1 = event.eventType;
                    arg2 = event.id;
                    arg3 = output;
                    await run(arg1, arg2, arg3)
                }
            } catch (error) {
                console.error('Error fetching row:', error.message);
            }
        }
    }
}
//=======================================================================================================
// main
(async () => {
    try {
        // TODO: Edit config.json to set desired sheet id and API token
        const config = require("./config.json");
        const PORT=process.env.PORT || 3000;
        initializeSmartsheetClient(config.smartsheetAccessToken, config.logLevel);

        // Sanity check: make sure we can access the sheet
        await probeSheet(config.sheetId);

        app.listen(PORT, () =>
            console.log("Node-webhook-sample app listening on port 3000"));

        await initializeHook(config.sheetId, config.webhookName, config.callbackUrl);
    } catch (err) {
        console.error(err);
    }
})();