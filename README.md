# BiralBond

বাংলাদেশের বিড়ালপ্রেমীদের প্ল্যাটফর্ম। Next.js (স্ট্যাটিক এক্সপোর্ট) + **Google Sheets (ডাটাবেস)** + Google Apps Script (ব্যাকএন্ড) + GitHub Pages।

**ডিপ্লয় করতে [DEPLOY.md](./DEPLOY.md) পড়ুন — ধাপে ধাপে বলা আছে।**

- `src/` — ওয়েবসাইট
- `google-apps-script/Code.gs` — ব্যাকএন্ড কোড (আপনার Google Sheet-এর Apps Script-এ বসাতে হয়) + শুরুর ডাটা
- `.github/workflows/deploy.yml` — অটো ডিপ্লয়

```
ভিজিটরের ব্রাউজার ──► GitHub Pages (স্ট্যাটিক সাইট)
        │
        └──► Google Apps Script (Web app) ──► Google Sheet (ডাটাবেস) + Google Drive (ছবি) + Gemini (Miau AI)
```
