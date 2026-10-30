# Hammerfall Website Demo

A clickable demo of the new public Vleissentraal website (auction calendar, auctions and lots, services, branches and their
teams, statistics, contact). It runs on sample data only, nothing is sent to a server.

## Open it
It must be served over http, opening `index.html` by double-clicking will not work.

1. Open a terminal in this folder.
2. Run `python -m http.server 8080`
3. Browse to `http://localhost:8080/`

The app itself is in the `demo` folder and also opens directly at `http://localhost:8080/demo/index.html`.

## Hosting
The whole folder can be uploaded as-is to any static host (GitHub Pages, Netlify, Azure Static Web Apps).
