# example_settings.html

A self-contained static demo of a settings page — a single HTML file with inline CSS and JavaScript, no build step and no dependencies beyond the Google Fonts it links. Open `example_settings.html` directly in a browser to view it.

## What it shows

- **Sidebar navigation** — a fixed left sidebar (Profile, Preferences, Notifications, Privacy) whose active item follows the URL hash; on narrow viewports it collapses to a horizontal scroll strip.
- **Toggle preferences** — accessible on/off toggles (`aria-pressed`, Space/Enter support) across the Preferences, Notifications, and Privacy sections. State persists in `localStorage` across reloads.
- **Profile card** — avatar, name, email, and an "Edit profile" button.
- **Dark mode** — the Dark mode toggle applies a `data-theme="dark"` theme and the choice is remembered on the next load.

It is a design/markup reference kept alongside the other static demos in this folder (`countdown_timer.html`, `demo.html`) — a starting point to copy from, not a wired-up feature of any automation script.
