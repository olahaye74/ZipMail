# 🧩 ZipMail Developer Guide

This document explains how to **set up a full development environment** for the ZipMail Outlook Add-in using the **Microsoft 365 Developer Program**.  
It allows you to sideload and debug the add-in without requiring a corporate Microsoft 365 or on-premise Exchange server.

---

## 1️⃣ Create a Microsoft 365 Developer Environment

### Step 1.1 — Register for a free developer tenant
Go to:

👉 [https://developer.microsoft.com/en-us/microsoft-365/dev-program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)

Click **“Join now”**, and sign in with a personal Microsoft account (no corporate account needed).

### Step 1.2 — Choose your setup
When prompted, select:
- **“Instant sandbox”** (recommended)
- Usage purpose: **“Develop and test Office add-ins”**

This will automatically create a Microsoft 365 E5 Developer tenant, including:
- Exchange Online (Outlook)
- SharePoint, Teams, and other M365 apps
- 25 test user accounts
- 1 admin account, for example:  
  `admin@yourtenant.onmicrosoft.com`

Once completed, you’ll receive:
- The admin username and temporary password  
- The tenant domain (e.g. `yourtenant.onmicrosoft.com`)

---

## 2️⃣ Activate Outlook for the Sandbox Tenant

1. Visit: [https://outlook.office.com/](https://outlook.office.com/)  
2. Log in with your **admin@yourtenant.onmicrosoft.com** account.  
3. You’ll land on a clean Outlook Web inbox — this is your **sandbox mail environment**.

✅ This Outlook Web environment **allows add-in sideloading**.

---

## 3️⃣ Run ZipMail Locally

In your ZipMail project directory, start the local development server:

```bash
npm run dev-server
```

You should see:

```
The dev server is running on port 3000
```

Your add-in files will now be served at:

```
[https://localhost:3000][https://localhost:3000]
```

## 4️⃣ Sideload the Add-in in Outlook Web

1. In Outlook Web, open the settings menu (⚙️ icon).
2. Click View all Outlook settings → Mail → Customize actions.
3. Scroll down and click View add-ins.
4. Choose Upload My Add-in → Add from file.
5. Select your local `manifest.xml` file.
Outlook will confirm:
> “Your add-in has been added successfully.”
✅ The ZipMail button should now appear in the ribbon when composing or reading an email.

## 5️⃣ Debug ZipMail

When you send or open an email using ZipMail:  
Outlook loads the add-in directly from [https://localhost:3000][https://localhost:3000]  
You can open the browser console (Cmd + Option + I) to inspect logs  
Webpack automatically hot-reloads changes as you edit your code  
You can edit files like:
- `/src/commands/commands.js`
- `/assets/ZipMailMessage.html`
- `/src/taskpane/password.html`
…and see changes live without redeploying.

## 6️⃣ Stopping and Cleaning Up

To stop the development server:  
`Ctrl + C`  
To remove the add-in from Outlook Web:  
Go to Manage Add-ins  
Delete ZipMail  

## 7️⃣ (Optional) Connect to Outlook Desktop

You can also add your developer tenant account to Outlook Desktop (if your IT policies allow external accounts):
Outlook → Preferences → Accounts → Add Email Account
Use admin@yourtenant.onmicrosoft.com
Outlook will detect your Exchange Online sandbox automatically
The ZipMail button will appear in your desktop Outlook as well
