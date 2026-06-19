# Frequency Visualizer for Microsoft Word 🌊

A sleek, interactive Microsoft Word Web Add-in that visualizes audio frequencies using an HTML5 Canvas animation. 

Originally built for documents related to sound therapy, binaural beats, and audio engineering, this tool runs directly in the Word task pane, allowing users to interact with frequency visualizations while reading or editing the document.

## ✨ Features

- **🎛️ Interactive Slider:** Smoothly adjust the frequency from 20 Hz up to 20,000 Hz.
- **🌊 Dynamic Wave Animation:** A custom HTML5 Canvas sine wave that visually scales its density and speed based on the selected frequency.
- **⚡ Quick Select Buttons:** Instantly jump to popular and therapeutic frequencies:
  - `10 Hz` (Alpha Waves / Binaural Beats)
  - `369 Hz` (Tesla / Perfect Balance)
  - `432 Hz` (Earth Tone)
  - `528 Hz` (Healing / Love)
- **📖 Dynamic Descriptions:** Automatically updates the text context to explain the physical or therapeutic meaning of the current frequency range.
- **🚀 Auto-Open:** Contains logic to automatically launch the task pane whenever the associated Word document is opened.

## 🛠️ Architecture

This is a modern Office Web Add-in. It does not use legacy COM or VSTO technologies. Instead, it consists of:
1. A **Web App** (`index.html`, `taskpane.js`, CSS) hosted via GitHub Pages.
2. A **Manifest File** (`manifest.xml`) that tells Microsoft Word where to find the web app and how to display it.

## 🚀 How to Install and Use (Sideloading)

Since this Add-in is not published in the official Microsoft AppSource store, you can easily install it locally on your Windows PC using a method called "Sideloading":

### Step 1: Prepare the Manifest
1. Download the `manifest.xml` file from this repository.
2. Create a folder on your computer (e.g., `C:\OfficeManifests`) and place the `manifest.xml` inside it.
3. Right-click the folder > **Properties** > **Sharing** tab > **Share...**. Add yourself, click **Share**, and copy the resulting Network Path (e.g., `\\YOUR-PC-NAME\OfficeManifests`).

### Step 2: Add to Microsoft Word
1. Open Microsoft Word and open a blank document.
2. Go to **File** > **Options** > **Trust Center** > **Trust Center Settings...**
3. Select **Trusted Add-in Catalogs** from the left menu.
4. In the **Catalog Url** box, paste the network path you copied in Step 1 and click **Add catalog**.
5. Check the box that says **Show in Menu** and click **OK**.
6. Restart Microsoft Word.

### Step 3: Launch the Add-in
Depending on your version of Microsoft Word, the Add-ins button might be located in different places:
- **Modern Word (Microsoft 365):** Go to the **Home** tab and look for the **Add-ins** button on the far right.
- **Classic Word:** Go to the **Insert** tab and click **My Add-ins**.
- **Developer Method:** Go to the **Developer** tab and click **Add-ins**.

Once you open the Add-ins menu:
1. Click on the **Shared Folder** tab at the top.
2. Select **FrequencyVisualizer** and click **Add**.
3. The interactive frequency tool will now open in the right task pane!
## 💻 For Developers

If you want to fork this project and develop it further:
1. Clone the repository.
2. Ensure you have Node.js installed.
3. Run `npm install` to install dependencies.
4. Run `npm start` to launch the local webpack dev server and automatically open Word for debugging.

## 📄 License

This project is licensed under the MIT License. You are free to use, modify, and distribute this software as you see fit. See the `LICENSE` file for more details.
