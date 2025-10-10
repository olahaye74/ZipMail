📦 ZipMail – Outlook Add-in for Mac / Windows
ZipMail is an Outlook add-in that automatically
compresses the message body and attachments into a single msg.zip,
optionally encrypts the ZIP with a password (AES),
automatically decompresses received messages so users see the original mail.

🚀 Quick Start
1️⃣ Clone the repository
git clone https://github.com/<your_username>/ZipMail.git  
cd ZipMail  

2️⃣ Run the setup script
npm run setup  
This command:
installs npm dependencies
installs the local HTTPS development certificates
creates a dist/ folder if missing
prints the next steps to start the dev server
If you’re on macOS, it also shows how to trust the local certificate manually.

🧩 Start the development server
npm run dev-server  
The add-in will be served at https://localhost:3000/

🧭 Load the add-in in Outlook
Once the dev server is running:
npm start
Then open Outlook → Preferences → Add-ins → Load add-in manually and select manifest.xml.
To stop debugging:
npm stop

🧰 Available npm scripts
Command	Description
npm run setup	One-time setup (install deps, certs, dist)
npm run dev-server	Launches Webpack dev server on https://localhost:3000
npm run build	Production build (dist/)
npm start	Starts Outlook with the add-in loaded
npm stop	Stops the add-in debug session
npm run lint	Runs ESLint checks
npm run lint:fix	Auto-fixes lint issues

⚙️ Requirements
Node.js ≥ 18 (Apple Silicon compatible)
Outlook 2024 LTS for Mac or Windows (modern experience)
Office.js, JSZip, Webpack 5, Babel, ESLint 9, Prettier

🧱 Project structure
ZipMail/
├── assets/                 # Static assets and icons
├── src/
│   ├── commands/           # Outlook command scripts (buttons, menus)
│   └── taskpane/           # Task pane HTML and JS
├── dist/                   # Generated build output (not versioned)
├── manifest.xml            # Add-in manifest for Outlook
├── package.json            # NPM scripts and dependencies
├── webpack.config.js       # Webpack configuration (ESM)
└── eslint.config.js        # ESLint + Prettier configuration  

🧹 .gitignore recommendation
node_modules/
dist/
build/
*.pem
*.crt
*.key
.DS_Store
npm-debug.log*
.env
coverage/
.vscode/*
!.vscode/settings.json
!.vscode/extensions.json  

🪪 License
This project is distributed under the MIT License.
© 2025 — ZipMail contributors.

💡 Notes
assets/ZipMailMessage.html contains the HTML template inserted into the email body when sending a zipped message.
During development, manifest.xml should point to https://localhost:3000/... URLs.
Keep package-lock.json under version control for reproducible builds.
On macOS, if Outlook cannot connect to https://localhost:3000/, trust the certificate manually once:
  sudo security add-trusted-cert -d -r trustRoot \
  -k /Library/Keychains/System.keychain \
  ~/.office-addin-dev-certs/localhost.crt  

🧑 Contributing
Contributions are welcome 🎉
Run npm run lint:fix before committing.
Ensure the add-in still works with npm run dev-server.
Keep commits small and clear.
Happy coding with ZipMail 🚀
