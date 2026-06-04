# Oil Well Production Analysis Tool

This repository contains a browser-based, interactive tool for analyzing oil well production data with log-log visualization and curve matching capabilities.

## Hosting Goal
The main goal of this repository is to host this application directly on GitHub Pages (`github.io`).

Because the application is built entirely using static client-side web technologies, it does not require a backend server and can be served directly by GitHub Pages.

## Project Structure
The app is modularized into three core files:
* `index.html` - Defines the semantic HTML structure, links the stylesheet and JavaScript logic, and loads external dependencies (Chart.js and Google Fonts) via CDN.
* `style.css` - Custom CSS styling featuring responsive grid layouts, glassmorphism panel styles, custom sliders, and subtle glow decorative elements.
* `app.js` - Client-side JavaScript containing data parsing (TSV/CSV inputs), piecewise regression calculations, outlier sensitivity filtering, and Chart.js configuration.

## Deployment Instructions
To deploy this tool to GitHub Pages:
1. Push this repository to a public or private GitHub repository.
2. Navigate to the repository on GitHub.
3. Click **Settings** (gear icon) -> **Pages** in the left sidebar.
4. Under **Build and deployment**:
   * Set **Source** to **Deploy from a branch**.
   * Under **Branch**, select `main` (or your default branch) and the `/ (root)` folder.
   * Click **Save**.
5. Once built by GitHub Actions, the application will be hosted at `https://<your-username>.github.io/<repository-name>/`.
