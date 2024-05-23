# PowerPoint Update Section Labels

Automatically update in-slide section labels.

## Usage

- Creates a section header for all slides with the "Footer" enabled (See:
  **Insert > Header & Footer**)
- Sections starting with `_` or `Default Section` will be ignored.
- Preserves the original formatting of the footer text box.
- Updates are triggered on new slides or when slides are re-arranged.
- The current section is colored in black on slides with light backgrounds and
  white on slides with dark backgrounds.

### Caveats

- If the entire text box is the "current section" color, (such as if there is
  only one initial section), when the text box is updated with new section text,
  the text will inherit that font color, so the text will no longer stand out.
  Click on the text box and click **Clear All Formatting** (next to font size
  adjustment) to revert the textbox to its original formatting. The next update
  will fix the highlighted text.
- It is possible in some cases that when slides in the first section are
  updated, the "selected section" color will propagate to the whole label. See
  the note above for how to resolve.
- The "slide change" update is technically processed when the _selected_ slide
  changes (due to PowerPoint limitations), so if you move the slide you have
  currently selected, the update will not process until you select a different
  slide.
- Moving slides between sections but maintaining the overall slide order will
  not trigger an update, however, they will be updated on the next update.

## Installation

1. Add `SectionStatusBar.ppam` to your `%APPDATA%\Microsoft\AddIns` folder.
2. In PowerPoint, navigate to **Developer > PowerPoint Add-Ins** OR go to
   **Settings > Add-Ins > Manage**, select "PowerPoint Add-ins" and click
   **Go**.
3. Click **Add New** and select `SectionStatusBar.ppam`. If you are prompted
   about enabling macros, click "Yes".
4. (Optional) Go to **Options > Quick Access Toolbar**, select "All Commands",
   scroll down to "Update Section Labels" and add it to your toolbar.

> Note: The add-in is loaded based on the file path, so if you move or delete
> `SectionStatusBar.ppam`, PowerPoint will not be able to load it.

To stop the functionality, the add-in can be disabled or removed at any time.
