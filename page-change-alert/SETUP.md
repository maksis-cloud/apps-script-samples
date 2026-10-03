# Page watcher — setup guide (about 5 minutes)

1. Open a new Google Sheet (sheets.new).
2. Menu Extensions -> Apps Script. Delete the sample code, paste the contents of Code.gs, click Save.
3. Go back to the sheet and reload the page. A new menu "Page watcher" appears.
4. Page watcher -> 1. Set up / restart. Google asks for permission:
   choose your account -> Advanced -> Go to (project name) -> Allow.
   The warning appears because the script is yours, not a published app. It only reads the pages you list and sends your alerts.
5. Three tabs are created: Settings, Watch, Log.

## Alerts on Telegram
1. In Telegram open @BotFather, send /newbot, choose a name. Copy the token it gives you.
2. Open your new bot and send it any message (for example "hi").
3. In a browser open https://api.telegram.org/bot<TOKEN>/getUpdates (put your token in place of <TOKEN>).
   Find "chat":{"id": 123456789 ... — that number is your chat ID.
4. In the Settings tab fill TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID. Several people: separate chat IDs with commas.

## Alerts by email
Fill ALERT_EMAIL in the Settings tab. You can use Telegram and email together.

## What to watch (Watch tab, one row per page)
Active  | TRUE to watch, FALSE to pause
Label   | short name shown in the alert
URL     | full link to a public page
Rule    | contains / not_contains / number_below / number_above / changed
Value   | the text or the number limit
Pattern | optional. Number rules: where the number is, e.g. Price:\s*\$([\d.,]+). Changed rule: which part of the page to watch.

Examples
TRUE | Back in stock | https://... | contains     | In stock
TRUE | Tickets       | https://... | not_contains | Sold out
TRUE | Laptop deal   | https://... | number_below | 1000
TRUE | New listings  | https://... | changed      |          | Results: \d+

## Check it works
Page watcher -> 2. Send test alert. Then 3. Check now. The Last status column shows what the watcher sees.
How often it checks: CHECK_EVERY_MINUTES in Settings (1, 5, 10, 15 or 30), then run 1. Set up / restart again.
Switch off: Page watcher -> Stop watching.

## Good to know
- One alert when the condition becomes true, not every check. It alerts again only after the condition goes away and comes back.
- If a page cannot be read 3 times in a row, you get one warning; watching continues.
- Works with public pages whose content is in the page HTML. Pages that need a login, show a captcha or load everything with JavaScript cannot be watched this way.
