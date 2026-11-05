# Hammerfall Website Demo

A clickable demo of the new public Vleissentraal website (auction calendar, auctions and lots, services, branches and their
teams, statistics, contact). It runs on sample data only, nothing is sent to a server.

## Open it
Double-click **Start Demo.bat**. It starts the demo and opens it in your browser at http://localhost:8090/.
Keep the black window open while you use the demo, and close it to stop.

It needs Python installed (https://www.python.org/downloads/). Opening `index.html` by double-clicking will not work:
the browser blocks the demo from loading its own files unless it is served like this.

## Without the launcher
1. Open a terminal in the `demo` folder.
2. Run `python -m http.server 8090`
3. Browse to `http://localhost:8090/`

## Hosting
The `demo` folder can be uploaded as-is to any static host (GitHub Pages, Netlify, Azure Static Web Apps).
