# ArNote

Simple notes that live in **your own Google Drive**: each note is a Google Doc, tags and the note list live in a Google Sheet, and attachments sit in a Drive folder. ArNote is an installable web app (phone or computer) with an Inbox for new notes, a Cabinet for filed ones, nested tags, search, batch actions, and a Drop folder that turns any file into a note.

**App:** https://mtnst8.github.io/ArNote/

## Using ArNote

1. Open the app link above on your phone or computer.
2. Paste your **pairing link** (or enter your server address and key) and tap **Connect**.
3. Install it:
   - **Android (Chrome):** ⋮ → *Install app* / *Add to Home screen*. Once installed, you can share photos and files from other apps straight into ArNote (Share → ArNote).
   - **iPhone (Safari):** Share → *Add to Home Screen*.
   - **Computer (Chrome/Edge):** the install icon in the address bar.

To add another device, open **Settings → Copy pairing link** in ArNote and open that link on the new device. Anyone with the pairing link can open your notes, so keep it private. To cut off every device, run `apiRotateKey` in your server's script editor and pair again.

## Set up your own ArNote

Each person runs their own small, free "server" in their own Google account, so nobody else can see your notes, including the person who made the app.

1. Create a new Google Sheet (it becomes your note index) → **Extensions → Apps Script**.
2. In the script editor, create files named `Code.gs`, `Batch.gs` and `Index.html`, and paste in the matching files from the [`server/`](server/) folder.
3. Save, pick the function **`setup`** and click **Run**. Approve the permission prompt (Google will say the app is unverified, because it's your own script: *Advanced → Go to …*). This creates an `ArNote` folder in your Drive with `Notes` and `Attachments` and a starter tag list.
4. Run **`installDropTrigger`**. This creates `ArNote/Drop`, where any file you put becomes an Inbox note within 15 minutes.
5. Run **`apiSetup`**. The log shows your **private key**.
6. **Deploy → New deployment → Web app**, with *Execute as: Me* and *Who has access: Anyone*. Copy the **Web app URL**.
7. Open the app, choose **Or enter server address and key**, and paste the URL and key. Then use **Settings → Copy pairing link** to set up your other devices.

"Who has access: Anyone" only lets the app reach your server. Every request must carry your private key, and anything without it is refused.

### Importing from Evernote

`server/Code.gs` includes an Evernote importer (Sheet menu **ArNote → Start / resume Evernote import**). It expects an `ArNote-import` folder in Drive with `notes.jsonl` and an `attachments/` folder, produced from your `.enex` export.

## Linked Drive files

Keep files where they are in Drive and still have them in ArNote:

- **Inbox → 🔗 Link Drive files**: browse My Drive or Shared with me (or search by name), tick files or folders, or paste links.
- Or in Drive (or the Drive phone app): **Organize → Add shortcut → ArNote › Link**. They show up within 15 minutes.
- A linked **folder** puts every file in it (and its sub-folders) in the Cabinet, tagged with the folder's name. New files are added on each check; stop it in **Settings → Linked Drive folders**.

ArNote never renames, edits, stars or deletes a linked file. "Remove link" only removes the note.

## Other Google accounts

Linked files are tagged with the Google account that owns them (for example `@brian-mountainstatebrewing`), under the tag **Accounts**. To bring in files from a second account, such as a work account, share its folders from that account with the Google account your ArNote runs in. Then link them from **🔗 Link Drive files → Shared with me**. Rename account tags like any other tag. **Settings → Google accounts** lists them.

## Sharing a tag with other people

Open a tag and choose **Share…** (or ⋯ → Share… on the Tags page). Type who it's for and **Create link**, then send them the link.

- People with the link see that tag's notes, including sub-tags, read-only. They can search, read, print and download attachments. They don't need a Google account or their own ArNote.
- Notes you add to the tag later show up for them automatically.
- Give each person or group their own link, so you can **Switch off** or delete one without affecting the others. **Settings → Shared links** lists them all.
- Renaming the tag keeps its links working. Deleting the tag switches its links off.
- A share link only reaches the notes in that tag. It can't change anything or see the rest of your notes.

## Printing

Select notes and choose **🖨 Print** (or Print on a single note). Google Docs are combined into one PDF; other files come as their own PDFs.

## How it's built

- `index.html`: the whole app (HTML/CSS/JS, no build step). It talks to your server with `fetch` (`{k: key, fn, a: args}`).
- `sw.js`: service worker, which makes the app installable, caches it for fast loading, and receives the Android share sheet.
- `manifest.webmanifest` and `icons/`: install metadata.
- `server/`: the Google Apps Script backend (`doPost` API in `Batch.gs`).

## Planned

- One shared server for several people, plus two-way sharing between ArNote users.
