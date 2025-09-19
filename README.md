# Outlook Add-in with React, TypeScript, and Microsoft Graph

This project is an **Outlook Add-in** built using **React** and **TypeScript**.  
It leverages the **Microsoft Authentication Library (MSAL)** to acquire user access tokens, enabling seamless integration with **Microsoft Graph APIs**.  

With this add-in, users can securely interact with Microsoft Graph services (such as Mail, Calendar, and OneDrive) directly within their Outlook environment.

---

## ✨ Features
- 🔑 Secure authentication using **MSAL.js**  
- 📧 Access to Microsoft Graph APIs from within Outlook  
- ⚛️ Modern **React + TypeScript** stack  
- 🛠️ Debug and run in both Outlook desktop and Outlook on the web  

---

## 🚀 Getting Started

### Prerequisites
- [Node.js](https://nodejs.org/) (V 18.20 LTS recommended)  
- [npm](https://www.npmjs.com/) (v 10.8.2)  
- Outlook desktop app **or** a Microsoft 365 account to use [Outlook on the web](https://outlook.office.com/mail/)

### SharePoint Connectivity
To connect it to SharePoint you need to register an app in Azure. To do so, go to the azure portl and select **App Registration** and simply register and app. 
Then open **MSAL.ts** file `(which is inside the auth folder)` and replace the Client ID of your registered app with the one in file. 

### Installation and run
1. Clone the repository:
   ```bash
   git clone https://github.com/shmata/Outlook_Add-in.git
   cd Outlook_Add-in
   code .
   // Open the msal.ts file and enter your client secret and tenant id
   npm install
   npm run start
   ```


### 🖥️ Usage

When you run npm run start, Outlook will launch automatically with the add-in sideloaded.

Open any email or compose a new one, then find your add-in under More Apps.

If you don’t have the desktop Outlook client installed, you can use the add-in in Outlook on the web:
👉 https://outlook.office.com/mail/

### 🛡️ Authentication & Microsoft Graph

This add-in uses MSAL to sign in users and acquire access tokens.
With these tokens, the add-in can call Microsoft Graph APIs to retrieve user data or perform actions on behalf of the user.

For more details, see:

Microsoft Graph API

MSAL.js Documentation

### 🧑‍💻 Contributing

Contributions are welcome! Feel free to fork the repo, open issues, or submit pull requests. ( Give it a star and become 0.1% cooler instatly ;) 
