# How to publish this to GitHub

## Option A — web browser (no tools needed)
1. Go to https://github.com/new, name the repository (for example `nigeria-ev-charging-2060`), choose Public, and
   create it **without** a README (this folder already has one).
2. On the empty repository page choose "uploading an existing file".
3. Unzip this package and drag the **contents** of the folder (`code/`, `raw_data/`, `figures/`, `workbook/`, `outputs/`,
   `README.md`, `LICENSE`, `CITATION.cff`, `.gitignore`, `requirements*.txt`) into the browser.
4. Commit. Your link will be `https://github.com/<your-username>/nigeria-ev-charging-2060`.

## Option B — command line
```bash
cd path/to/this/folder
git init -b main
git add .
git commit -m "Model, data and code for Nigeria EV charging infrastructure to 2060"
git remote add origin https://github.com/<your-username>/nigeria-ev-charging-2060.git
git push -u origin main
```

## After publishing
- Replace `USERNAME/REPONAME` in `CITATION.cff` with the real path.
- For a citable DOI (useful for the journal's data-availability statement), link the repository to Zenodo
  (https://zenodo.org/account/settings/github/) and publish a release; Zenodo then issues a DOI for that release.
