# Regen Inverter Prices

Inverter price list for Regen Renewables (Pvt) Ltd.

- **Procurement** updates prices, supplier, stock and validity. They can also add new brands and models and hide discontinued ones.
- **All other teams** (Sales, Design, Management, Finance, Operations) see the same list live, read only.
- Every change is saved in the **History** tab of the Google Sheet. It also shows in the app under **Recent changes** and in each model's price history.
- The first list is the 202 inverters from the SLD tool (`SLD-Genarator/inverter-data.js`). **Sync from SLD tool** adds any models you add to the SLD tool later. It never changes existing prices.

App link (after setup): `https://udithapana.github.io/Regen-Inverter-Prices/`

## 1. Google Sheet (server), one time

1. Create a new Google Sheet named **Regen Inverter Prices**.
2. Open **Extensions → Apps Script**. Delete the sample code, paste all of `Code.gs`, then press **Save**.
3. In the function list at the top, choose **setup** and press **Run**. Allow the permissions it asks for.
   This creates three tabs:
   - **Inverters**: the 202 SLD models
   - **History**
   - **Settings**: the two team keys and the price note
4. Choose **Deploy → New deployment → ⚙ Web app** and set:
   - Execute as: **Me**
   - Who has access: **Anyone**

   Press **Deploy**, then copy the link that ends in `/exec`.

### Team keys (Settings tab)
| Setting | Who gets it | What it allows |
|---|---|---|
| Procurement key | Procurement team only | Edit prices, add brands and models, hide models, edit the price note, sync from the SLD tool |
| View key | Everyone else | View only |

You can change a key in the Settings tab at any time. Anyone using the old key is signed out and asked for the new one.

### If you change Code.gs later
Choose **Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy**. This keeps the same `/exec` link.

## 2. GitHub Pages

1. Create a new **public** repository named `Regen-Inverter-Prices` under `udithapana`.
2. Upload every file from this folder except `Code.gs`. You can upload `Code.gs` too as a backup; the app does not use it from GitHub.
3. Open **Settings → Pages → Build and deployment**. Choose **Deploy from a branch**, then branch `main`, folder `/ (root)`, and **Save**.
4. Optional: in `index.html`, paste the `/exec` link into `API_URL_DEFAULT = ''`. If you do, people don't need to paste the server link when they sign in.

## 3. Using it
- **Sign in:** enter your name, team and team key. The app remembers you on that device.
- **Find a model:** search by brand, model, kW or supplier. You can also filter by brand, type (On-grid / Hybrid / BESS) or phase, and sort by price.
- **Badges:**
  - **NEW PRICE:** the price changed in the last 7 days.
  - **OLD PRICE:** the price hasn't been updated for 90 days.
  - **VALIDITY OVER:** the "valid until" date has passed.
  - **▲ / ▼ from Rs …:** shows the last price change.
- **Procurement:**
  - Tap a model to edit it.
  - **Bulk price edit:** type many prices, then press **Save all**.
  - **＋ Add inverter:** to add a new brand, type its name in the Brand box.
  - **Hide (discontinued):** removes a model from everyone's list but keeps its history.
  - **Sync from SLD tool:** reads `../SLD-Genarator/inverter-data.js` and `index.html` on this site. If that fails, you can choose the file from your computer.
- **Exports:** **Excel (CSV)** and **Print / PDF** export the list as filtered. Both carry the Regen copyright and confidentiality notice.
- If two Procurement users edit the same model at the same time, the second save is stopped and asked to reload. This way nobody overwrites the other's price by mistake.

Edit prices in the app rather than directly in the Sheet. Changes made in the Sheet are not written to History.

## Notes for these apps on udithapana.github.io
- `manifest.webmanifest` has its own `"id": "/Regen-Inverter-Prices/"` and `"scope": "./"`. The maskable icon is a separate file.
- The service worker uses cache names starting with `regen-inverter-prices-` and only deletes its own caches.
- localStorage keys all start with `rip_`. The app never calls `localStorage.clear()`.

© Regen Renewables (Pvt) Ltd. Confidential — internal use only.
