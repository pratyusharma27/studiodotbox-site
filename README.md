# Studio Dotbox Website

A static website for Studio Dotbox. Dark theme. Built in plain HTML, CSS, and JavaScript. No build step required to view, but a small Python script regenerates pages when you update content.

## How content works on this site

The Projects and Journal articles for the website are managed in **studiodotbox-content.xlsx**. You fill in the Excel file. You drop matching images into the right folders. You run `generate_pages.py`. The website updates.

This means:

- Adding a new project = adding a row to the Projects sheet + dropping images in `images/projects/{slug}/` + running the script.
- Publishing a new journal article = adding a row to the Journal sheet + dropping images in `images/journal/` + running the script.
- No HTML editing required for content updates.

The static pages (Home, Studio, Services, How we work, Philosophy, Contact, Work with us, 404) are also rebuilt by the script every time it runs, so any future updates to those pages flow through the script.

## What's in this folder

```
studiodotbox-site/
├── index.html              Home page (with signature animation)
├── studio.html             About the studio and Pratyush
├── services.html           Six disciplines
├── how-we-work.html        Five-stage process
├── philosophy.html         Design Out of the Box
├── work.html               Project archive (filterable)
├── work/                   Individual project pages (auto-generated)
├── journal.html            Journal archive
├── journal/                Individual journal articles (auto-generated)
├── contact.html            Eight-field enquiry form
├── work-with-us.html       Three anchored sections
├── 404.html                Page-not-found page
├── css/                    Stylesheets (main, pages, home, animations)
├── js/                     JavaScript (main, animations)
├── images/                 All images (see image guide below)
├── studiodotbox-content.xlsx  Content workbook (Projects + Journal)
├── generate_pages.py       Page generator script
├── vercel.json             Vercel deployment config
└── README.md               This file
```

## Setting up

You only need to do this once.

### 1. Install Python (if you don't have it)

Mac/Linux: usually pre-installed. Check with `python3 --version`.

Windows: download from [python.org](https://python.org). During install, check "Add Python to PATH."

### 2. Install the openpyxl library

The generator script reads the Excel file using a Python library called openpyxl.

Open Terminal (Mac/Linux) or Command Prompt (Windows) and run:

```bash
pip3 install openpyxl
```

That's it for setup.

## Three things to do before going live

### 1. Replace the logo

The current logo is a placeholder SVG. Replace it with your real logo file:

1. Save your logo as `images/logo.svg` (recommended) or `images/logo.png`.
2. The logo should be the wordmark in white (because the site is on a black background).

If you want, the existing red logo can be used too. Just make sure the colour reads against black.

### 2. Add your Web3Forms access key

The contact form and the subscribe form need a Web3Forms key to actually send.

1. Go to [web3forms.com](https://web3forms.com).
2. Enter `studiodotbox@gmail.com` and click "Create access key."
3. Copy the key from the email Web3Forms sends.
4. Open Terminal, navigate to the studiodotbox-site folder, and run:

```bash
# Mac/Linux:
find . -name "*.html" -exec sed -i '' 's/YOUR_WEB3FORMS_KEY_HERE/your-actual-key-here/g' {} +

# Linux (no quotes after -i):
find . -name "*.html" -exec sed -i 's/YOUR_WEB3FORMS_KEY_HERE/your-actual-key-here/g' {} +
```

Replace `your-actual-key-here` with the real key. Also update the same key inside `generate_pages.py` so future regenerations include it.

### 3. Add the hero image and portrait

Two key images for the static pages:

- `images/hero.jpg` — the large image at the top of Home. **2400 × 1400 pixels.**
- `images/pratyush-portrait.jpg` — your portrait on the Studio page. **1600 × 900 pixels.**

Drop them in the `images/` folder with those exact filenames.

## How to add new projects and articles

### Adding a project

1. Open `studiodotbox-content.xlsx`.
2. Go to the **Projects** sheet.
3. Add a new row. Fill in every column (slug, name, sector, stage, year, location, description, narrative, number of gallery images, and optional fields).
4. Save and close the Excel file.
5. Create a folder for the project at `images/projects/{slug}/` (e.g. `images/projects/villa-noida/`).
6. Drop your images in that folder, named exactly as follows:
   - `hero.jpg` — the project hero image. **2400 × 1350 pixels.**
   - `thumb.jpg` — the project card thumbnail (used on Home and Work archive). **1200 × 900 pixels.**
   - `01.jpg`, `02.jpg`, `03.jpg` ... — gallery images, numbered sequentially. **1800 × 1200 pixels.** Number of images must match what you put in the "Number of gallery images" column in Excel.
7. Open Terminal, navigate to the studiodotbox-site folder, and run:

```bash
python3 generate_pages.py
```

8. Open `index.html` in your browser to preview the new project.

### Adding a journal article

1. Open `studiodotbox-content.xlsx`.
2. Go to the **Journal** sheet.
3. Add a new row. Fill in slug, title, category, reading time, publish date, excerpt, and the full body. Use a blank line between paragraphs in the body field.
4. Save and close the Excel file.
5. Drop two images in the `images/journal/` folder:
   - `{slug}.jpg` — the article hero image. **1800 × 1200 pixels.**
   - `{slug}-thumb.jpg` — the card thumbnail. **800 × 600 pixels.**

   For example, for an article with slug `the-questionnaire-method`, you'd save:
   - `images/journal/the-questionnaire-method.jpg`
   - `images/journal/the-questionnaire-method-thumb.jpg`
6. Run the generator:

```bash
python3 generate_pages.py
```

7. Preview by opening `index.html`.

### Featuring a project or article on Home

In the Excel file, set the "Featured on Home" column to `YES` for any project or article you want to appear on the Home page (in the Selected work strip or From the Journal section). Leave blank to hide.

The script picks up to five featured projects and up to two featured articles for Home.

## Image sizes summary

| Image | Where | Size |
|---|---|---|
| Hero | Top of Home | **2400 × 1400 px** |
| Portrait | Studio page | **1600 × 900 px** |
| Project hero | Top of project page | **2400 × 1350 px** |
| Project thumbnail | Cards on Home and Work | **1200 × 900 px** |
| Project gallery | Inside project page (numbered 01.jpg, 02.jpg ...) | **1800 × 1200 px** |
| Journal hero | Top of article | **1800 × 1200 px** |
| Journal thumbnail | Cards on Home and Journal archive | **800 × 600 px** |

### Image format and tips

- JPG for photographs.
- PNG for logos or graphics with transparency.
- Compress before uploading. [tinypng.com](https://tinypng.com) is free and good. Aim for each image to be under 500KB.
- Sizes above are recommendations. The site will scale images, but matching these sizes ensures crisp display without unnecessary file weight.

## Image folder structure

```
images/
├── logo.svg                          (logo)
├── hero.jpg                          (Home hero)
├── pratyush-portrait.jpg             (Studio page)
├── projects/
│   ├── villa-noida/
│   │   ├── hero.jpg
│   │   ├── thumb.jpg
│   │   ├── 01.jpg
│   │   ├── 02.jpg
│   │   └── ... (etc, up to the count in Excel)
│   └── brand-x-flagship/
│       └── ...
└── journal/
    ├── designing-at-every-scale.jpg
    └── designing-at-every-scale-thumb.jpg
```

## How to preview locally

After adding content and running `generate_pages.py`, double-click `index.html` and the site opens in your browser. Everything works: animations, mobile menu, form (in preview mode without the Web3Forms key).

## How to deploy to Vercel

The recommended path:

1. Create a free [GitHub](https://github.com) account.
2. Create a new repository called `studiodotbox-website`.
3. Upload the entire studiodotbox-site folder to the repository.
4. Sign up at [vercel.com](https://vercel.com), connect your GitHub.
5. Click "Add New" → "Project" → select your repository.
6. Vercel detects it as a static site. Click "Deploy."
7. Within 30 seconds, your site is live at `studiodotbox-website.vercel.app`.

For your own domain: in Vercel, go to the project Settings → Domains. Add `studiodotbox.com` and update your domain's DNS records with the values Vercel gives you. Live within an hour.

### Updating after deployment

After every Excel update + `generate_pages.py` run, push the updated files to GitHub. Vercel automatically redeploys within a minute.

## Getting help

If anything is unclear, return to the conversation that built this site. Or share the specific issue with a developer comfortable with HTML/Python and they will get you unstuck.

Studio Dotbox · studiodotbox@gmail.com · +91 9571 129 136
