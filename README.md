# Nursing Career Data

Astro static-site starter for nursingcareerdata.com.

## Included
- Astro static generation
- Responsive homepage and core information pages
- RN salary index and data-driven state route
- Canonical URLs and descriptions
- Cloudflare Workers Static Assets configuration
- Salary data schema with intentionally unverified/empty figures

## Data integrity
The included `src/data/rn-state-wages.json` is a schema example, not a real salary dataset. The state route only generates pages for records with a numeric `annualMeanWage` and a `referenceYear`. Do not publish invented or estimated values as BLS figures.

## Run locally (Windows PowerShell)
Install Node.js LTS and Git, then in this folder run:
```powershell
npm install
npm run dev
```
Open the local URL printed by Astro (usually http://localhost:4321).

Build:
```powershell
npm run build
```
Generated files are in `dist/`.

## Push to the existing GitHub repository
Open PowerShell in this folder:
```powershell
git init
git add .
git commit -m "Initial Nursing Career Data Astro site"
git branch -M main
git remote add origin https://github.com/MasimoMaserati/nursingcareerdata.git
git push -u origin main
```
If `origin` already exists, inspect with `git remote -v` rather than adding it again.

## Cloudflare
This starter includes `wrangler.jsonc` for Workers Static Assets. After dependencies are installed and Wrangler is authenticated (`npx wrangler login`), deploy with:
```powershell
npm run deploy
```
For Git-based deployment, connect the repository in Cloudflare Workers and configure install/build/deploy commands as appropriate. Add the custom domain through the Worker Domains & Routes settings. Do not create manual DNS records until Cloudflare's custom-domain flow tells you what is required.

## Next milestone
Import and validate the BLS OEWS state-level RN dataset, then add verified records to `src/data/rn-state-wages.json`. Expand to city/metro pages only after geography mappings are validated.
