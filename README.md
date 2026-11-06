# merchitecorp

Public website for **Merchitecorp — Spectrum and Promotion**.

Live site (GitHub Pages, once Pages is enabled on the `main` branch, root folder):

https://spectrumandpromotion-arch.github.io/merchitecorp/

Open it on any phone, on any network — it is the same URL everywhere.

## Pages

- `index.html` — home page: promotion / merchandising, services, contact (spectrumandpromotion@gmail.com)
- `login.html` — login page (email + password)
- `dashboard.html` — welcome page after login, with a log-out button
- `style.css` — shared, mobile-responsive styles

## About the login

This is a static site hosted on GitHub Pages, which has no server or backend,
so the login is a client-side demo only: submitting the login form saves a
session flag in that browser (`localStorage`) and opens the dashboard. No
password is checked, sent, or stored, and it is not real secure authentication.
Do not enter a password you use anywhere else.

## Publishing

Enable GitHub Pages in the repository Settings → Pages, source: Deploy from a
branch, branch `main`, folder `/ (root)`. The site then appears at the URL above.
