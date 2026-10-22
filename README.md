# Teams Transcript Extractor

A userscript that saves the full transcript of a Microsoft Teams meeting recording as a plain-text file.

The Teams recap page only renders the transcript lines currently on screen, so copy-pasting or "select all" gets you a few dozen lines at most. This script scrolls through the transcript, captures each line as it renders, and lets you download or copy the whole thing.

## Output

```
[Jane Doe started transcription]

John Smith  1:53
Hi, any update from the vendor?

Jane Doe  1:57
Still waiting on them, they haven't replied yet.
```

## Install (Tampermonkey / Violentmonkey)

1. Install [Tampermonkey](https://www.tampermonkey.net/) or [Violentmonkey](https://violentmonkey.github.io/).
2. Create a new script and paste in `teams-transcript-extractor.user.js`, or open the raw file URL to install it directly.
3. On recent Chrome versions, open `chrome://extensions` → Tampermonkey → Details and enable **Allow User Scripts** (or enable Developer mode on older versions), and set Site access to "On all sites".

## Usage

1. Open the meeting recap in Teams (web) and switch to the **Transcript** tab.
2. A small panel appears at the top of the transcript.
3. Click **Auto-scroll & capture**. The counter shows captured lines against the total.
4. Click **Download .txt** or **Copy**.

Keyboard shortcuts (click inside the transcript first):

- `Alt+Shift+S`: start/stop auto-scroll
- `Alt+Shift+D`: download

## Without an extension (DevTools console)

The transcript lives in a cross-origin iframe (`<tenant>-my.sharepoint.com/.../xplatplugins.aspx`), so the code must run inside that frame:

1. Open DevTools → Console.
2. In the frame dropdown (top left, defaults to `top`), select the `xplatplugins.aspx` frame.
3. Paste the contents of `teams-transcript-extractor.user.js` and press Enter.
4. Use the panel, or run:

   ```js
   await __transcriptExtractor.autoScroll();
   __transcriptExtractor.download();
   ```

## How it works

- Rows are keyed by their `data-list-index`, so re-renders and scrolling back up never duplicate lines.
- The total line count comes from `aria-setsize`, and any gaps are logged to the console.
- Lines rendered without a speaker header inherit the previous speaker.
- Auto-scroll repeatedly scrolls the last rendered row into view, which works regardless of which element is the actual scroll container.

## Notes

- This relies on the Teams/Stream transcript DOM (Fluent UI `ms-List` cells, `sub-entry-*` ids). Microsoft changes this markup from time to time, so it may need selector updates.
- Only works on transcripts you already have access to view. Nothing is sent anywhere; the file is generated locally in the browser.
- If you only need the transcript once, check the transcript's own **Download** menu first; it may already give you a `.docx` or `.vtt` depending on your permissions.

## License

MIT
