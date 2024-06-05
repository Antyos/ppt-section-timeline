# PowerPoint Section Timeline Add-In

Automatically update an in-slide section timeline.

![Slide Timeline Example](SectionTimelineExample.png)

> The section timeline is in the top right corner.

## Usage

- Adds a command group called "Section Timeline" in the Design Tab.
- Creates a section header for all slides with the "Footer" enabled (See:
  **Insert > Header & Footer**)
- Sections starting with `_` or `Default Section` will be ignored.
- Preserves the original formatting of the footer text box.
- Updates are triggered on new slides or when slides are re-arranged. This can
  be toggled on/off in the **Design > Section Status Bar > Auto Update Timeline**.
- The current section is colored in black on slides with light backgrounds and
  white on slides with dark backgrounds. The current section is bolded by
  default, but this can be toggled in **Design > Section Status Bar > Bold Active Section**.

### Caveats

> [!IMPORTANT]
>
> If you are adding sections to an existing presentation, make sure you have the
> `Auto Update Timeline` option ***disabled*** or have footers hidden while
> adding your first 2 sections so that the style for active sections is not
> propagated to inactive sections, i.e., adding the first section will make the
> entire footer have black text, so when you add a second section, it will
> assume you also have black text for all inactive sections. See below for how
> to resolve this.

- If the entire text box is the "current section" color, (such as if there is
  only one initial section), when the text box is updated with new section text,
  the text will inherit that font color, so the text will no longer stand out.
  Click on the text box and click **Clear All Formatting** (next to font size
  adjustment) to revert the textbox to its original formatting. The next update
  will fix the highlighted text.
- It is possible in some cases that when slides in the first section are
  updated, the "selected section" color will propagate to the whole label. See
  the note above for how to resolve.
- The "slide change" update is technically processed when the *selected* slide
  changes (due to PowerPoint limitations), so if you move the slide you have
  currently selected, the update will not process until you select a different
  slide.

## Extras

- **Change Video**: Swap a video with another video on disk while preserving the
  animations.
    - Can be found in the right-click context menu for videos or in the "Video
      Format" tab, i.e., **Video Format > Change > Change Video**.
    - Note: Using this on a video in a (non-nested) group will work, however
      undoing the replace may corrupt the animations on the slide.

## Installation

1. Add `SectionTimeline.ppam` to your `%APPDATA%\Microsoft\AddIns` folder (Windows).
2. In PowerPoint, navigate to **Developer > PowerPoint Add-Ins** OR go to
   **Settings > Add-Ins > Manage**, select "PowerPoint Add-ins" and click
   **Go**.
3. Click **Add New** and select `SectionTimeline.ppam`. If you are prompted
   about enabling macros, click "Yes".

> Note: The add-in is loaded based on the file path, so if you move or delete
> `SectionTimeline.ppam`, PowerPoint will not be able to load it.

To stop the functionality, the add-in can be disabled or removed at any time.
