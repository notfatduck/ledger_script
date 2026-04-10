# ledger_script

This is a Google Apps Script that turns a standard Google Sheet into an automated tracking system for recurring Zelle payments. It reads your email for Zelle receipts, matches the payment to a person in your spreadsheet, calculates how many months they just paid for, and checks off the corresponding boxes on a timeline. 

It also calculates balances and can ping people on Discord if they fall behind.

## Key Features

* **Hands-Off Tracking:** Scans Gmail hourly for new Zelle payment alerts and archives the emails once processed.
* **Smart Logging:** Takes the payment amount, divides it by the sender's monthly rate, and automatically checks off the correct number of months on your ledger.
* **Live Summary:** Creates a dashboard showing exactly who is up to date, who is ahead, and who is falling behind.
* **Automated Nudges:** Optionally sends a Discord message on the 22nd of every month tagging anyone who owes money.

## How to Set It Up

1. **Add the Code:** Open a new Google Sheet, click **Extensions > Apps Script**, and paste the code into the editor.
2. **Adjust the Settings:** Look at the `CONFIG` block at the very top of the code. The default `GMAIL_QUERY` looks for Chase Bank emails (`no.reply.alerts@chase.com`). If you use a different bank, you will need to update this query to match the sender address of your specific Zelle alerts.
3. **Run the Setup:** Save the code, then go back to your Google Sheet and refresh the page. A new menu called **Ledger Tools** will appear at the top. Click it and select **Setup Sheet**.
4. **Add Your People:** In the newly created `History` tab, enter the names and monthly rates of the people you are tracking. Make sure the names closely match how they appear in the Zelle emails.
5. **Turn on the Automation:** Go to the **Ledger Tools** menu again and click **Reset Triggers**. This tells Google to run the email scanner every hour in the background.

## Optional: Discord Reminders

If you want the script to automatically remind people to pay, you need to link a Discord Webhook.

1. In the Apps Script editor, click the gear icon on the left to open **Project Settings**.
2. Scroll down to **Script Properties** and click **Add script property**.
3. Enter `DISCORD_WEBHOOK_URL` in the Property field.
4. Paste your actual Discord webhook link into the Value field and save. 

## Important
* **Exact Math Required:** The script relies on clean math to work. If a person's rate is $20 a month, and they send $40, the script knows to check off two months. If they send $25, the script will reject it and log an error because it is not a clean multiple of their rate.
* **Memory Cleanup:** The script remembers which emails it has already processed so it does not double-count payments. It clears this memory every 30 days to prevent your Google account from running out of storage space.
* **Manual Adjustments:** If you ever need to manually adjust someone's timeline, you can just click the checkboxes in the `History` tab yourself. The `Summary` tab will automatically recalculate their balance based on the total number of checked boxes.
