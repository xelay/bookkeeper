# Bookkeeper — Double-Entry Personal Finance PWA

Offline-first personal finance web app with double-entry bookkeeping.

## Features

- **Double-entry bookkeeping** — every transaction has a Debit and Credit account
- **Custom accounts** — create wallets like "Cash USD", "IIS", "CNY Deposit", "Income: Salary"
- **Transaction journal** — searchable, filterable by year/month
- **Balance report** — account balances with drill-down ledger per account
- **Import/Export XLSX** — via SheetJS
- **IndexedDB storage** — reliable local persistence
- **PWA / iOS support** — works offline, installable on iOS home screen

## Setup

1. Copy all files to a static web server (nginx, GitHub Pages, Netlify, etc.)
2. Enable GitHub Pages: Settings → Pages → Branch: main → Save
3. App will be live at `https://xelay.github.io/bookkeeper/`

## File Structure

```
bookkeeper/
├── index.html       — Main app shell
├── app.js           — Business logic (IndexedDB, double-entry, SheetJS)
├── style.css        — Mobile-first design, dark/light theme
├── manifest.json    — PWA manifest
├── sw.js            — Service worker (offline caching)
└── icons/
    ├── icon-192.png
    └── icon-512.png
```

## Import Format (XLSX)

Sheet named **Transactions** with columns:
`date, debitAccount, creditAccount, amountRub, currency, foreignAmount, rate, description, accountTypeDebit, accountTypeCredit`
