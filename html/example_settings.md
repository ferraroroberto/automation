Build a settings page with a sidebar navigation, toggle preferences, and a profile section.

(user ui-design-brain skill)

Create settings page HTML structure with sidebar nav
Add toggle preferences section and profile section
Style with distinct typography, palette, and responsive layout
settings.html(new)
+539
<!DOCTYPE html><html lang="en"><head>  <meta charset="UTF-8">  <meta name="viewport" content="width=device-width, initial-scale=1.0">  <title>Settings</title>
settings.html
+23
      </section>
      </section>      <section class="section" id="privacy" aria-labelledby="privacy-heading">        <div class="section-header">          <h2 id="privacy-heading" class="section-title">Privacy</h2>          <p class="section-desc">Control what data is shared and stored.</p>
Adding dark-mode support so the Dark mode toggle applies the theme:

settings.html
+13
    [data-theme="dark"] {      --bg: #1a1a1a;      --bg-card: #252525;      --text: #f2f0eb;      --text-muted: #a39e93;
settings.html
+3
          if (key) savePref(key, !pressed);        });
          if (key) savePref(key, !pressed);          if (key === 'darkMode') {            document.documentElement.setAttribute('data-theme', !pressed ? 'dark' : 'light');          }        });
settings.html
+3
            btn.setAttribute('aria-pressed', prefs[key] ? 'true' : 'false');          }
            btn.setAttribute('aria-pressed', prefs[key] ? 'true' : 'false');            if (key === 'darkMode' && prefs[key]) {              document.documentElement.setAttribute('data-theme', 'dark');            }          }
settings.html
+4
      // Sidebar: set active page from hash
      if (loadPrefs().darkMode) {        document.documentElement.setAttribute('data-theme', 'dark');      }      // Sidebar: set active page from hash

Summary of what’s in place:

File: html/settings.html

1. Sidebar navigation

Fixed left sidebar (260px) with “Settings” title and 4 items: Profile, Preferences, Notifications, Privacy.
Each item has an icon and label; active item uses a left border and background.
Clicking a link updates the hash and smoothly scrolls to that section; active state follows the current hash.
On viewports ≤768px the sidebar becomes a horizontal scroll strip under the title, with the same links.
2. Toggle preferences

Preferences: Dark mode, Compact density, Reduced motion, Keyboard shortcuts.
Notifications: Email notifications, Push notifications.
Privacy: Usage analytics, Profile visibility.
Each row has a label and short description on the left and a toggle on the right.
Toggles use aria-pressed and keyboard support (Space/Enter).
State is stored in localStorage and restored on load.
Dark mode is wired: turning it on applies a data-theme="dark" theme (dark background, light text, adjusted borders/shadows).
3. Profile section

First section includes an avatar (initials “JD”), name “Jane Doe”, email, and an “Edit profile” button.
Styled as a card with a short description under the section title.
Design

Fonts: Instrument Serif for headings, Source Sans 3 for body (loaded from Google Fonts).
Colors: Warm off-white background (#f8f6f2), near-black text, accent #b8860b (gold) for active states and buttons.
Layout: 8px-based spacing, clear focus rings, touch-friendly toggle size (~44px wide).
Accessibility: Semantic sections and headings, aria-current and aria-pressed, focus styles, and section IDs for in-page navigation.
Open html/settings.html in a browser to use the page. Toggling “Dark mode” switches the theme; all toggle states persist across reloads.