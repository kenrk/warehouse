# 📦 Warehouse Project

The **Warehouse** project is a Google Apps Script-based application that combines HTML, CSS, and JavaScript to build a simple web app. Its main goal is to provide an interactive interface for managing a data warehouse online.

---

## 📂 File Structure

- **Code.gs**
- The main script is based on Google Apps Script.
- Handles backend logic, communication with Google Sheets/Database, and server-side functions.

- **JavaScript.html**
- Contains client-side JavaScript code.
- Manages user interaction, event handling, and communication with `Code.gs` via `google.script.run`.

- **Stylesheet.html**
- CSS file for styling the web app's appearance.
- Provides layout, colors, and design for a more consistent and user-friendly UI.

- **WebApp.html**
- The main HTML file.
- Serves as the entry point for web applications, combining CSS and JavaScript, and displaying the interface to the user.

---

## 🚀 How to Run

Google Sheet Link: https://docs.google.com/spreadsheets/d/1TXud--lMGc_9o9Dk1DcpTJyGrw_WD1cbZspV4Iz4Ows/edit?usp=sharing
Web App Link: https://script.google.com/macros/s/AKfycbx56mzLOeIuAqRHXM_ZLd9XoWeLPvwk6jGiVrfd9uINQMB0-rRnQJUVk7Z11Zm_88M-/exec

**If Web App cannot be opened, open it in incognito**

Or for new google sheets file, follow these steps:
1. Open the Google Apps Script editor.
2. Upload all files (`Code.gs`, `WebApp.html`, `JavaScript.html`, `Stylesheet.html`).
3. Ensure the `doGet(e)` function in `Code.gs` returns `HtmlService.createTemplateFromFile('WebApp')`.
4. Deploy as a **Web App**:
- Click **Deploy > New Deployment**.
- Select **Web App**.
- Set access as needed (for example: Anyone with the link).
5. Access the provided URL to use the application.

---

## ✨ Key Features

- Backend integration with Google Apps Script.
- Interactive web interface based on HTML, CSS, and JavaScript.
- Client-server communication uses `google.script.run`.
- Modular design: separation of logic, display, and interaction.

---

## 🛠️ Development

- Add new functions in `Code.gs` for CRUD (Create, Read, Update, Delete) operations.
- Use `JavaScript.html` to call backend functions and process data.
- Customize `Stylesheet.html` to match your company's branding.
- Extend `WebApp.html` to add new pages or components.

---

## 📌 Notes

- Make sure your browser cache is cleared when making CSS/JS changes.
- Use specific CSS selectors to avoid conflicts with other themes/plugins.
- Document each function in `Code.gs` for easy maintenance.
