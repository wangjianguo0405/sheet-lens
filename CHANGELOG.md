# Changelog

**English** | [简体中文](CHANGELOG.zh-CN.md) | [日本語](CHANGELOG.ja.md)

One section per release. The section for the version being published becomes
both the release notes on GitHub and the notes the app shows when it offers the
update — so write it for whoever is being asked to install it, not for whoever
wrote the code.

Within a section the entries are grouped: **Added** for what a sheet could not
do before, **Improved** for what it now does differently, and **Fixed** for what
it used to get wrong. A release with nothing to say under one of them leaves
that group out.

## [0.3.17] - 2026-10-16

A workbook's own hidden sheets are respected, and the file tabs grow the two things a full strip needs: a list of what is open, and a way to read a file again.

### Added

- The sheets a workbook hides are left out, as Excel leaves them out: no tab, nothing to click in the sheet list (it is listed there, greyed, and says so), and no search hits — a hidden sheet cannot be activated at all, and a hit is a place the cursor is about to go. A workbook opens on the first sheet it shows
- Hide and Unhide on a sheet tab's right-click menu, with the accelerator letters Excel names them by. Unhide opens the list of hidden sheets to pick one from; hiding the sheet on screen moves to the tab that takes its place, and the last sheet still shown cannot be hidden. Nothing is written back to the file: what the reader hides is theirs, and lasts as long as the workbook is open
- A chevron at the right end of the file tabs, for reaching a file that no longer fits across the strip. It appears the moment the tabs stop fitting — the same condition Chrome and Firefox put theirs under — and lists every open file in strip order, with the start page last; the one showing is ticked, and each line carries a × that closes it without closing the list
- A file tab can read its workbook again, from its right-click menu. That is for the file that changed underneath us, which Excel being open on it is the usual cause of. What the tab remembered *about the file* goes — the sheets that no longer exist, and a cursor or a current sheet pointing at one — while everything the reader put on it stays
- The menus keep a column at the left for an icon, as Excel's do, so the labels line up whether or not an entry has one. Hide and Unhide are a closed and an open eye, and reading a file again is a circular arrow

### Improved

- Opening a workbook, or pressing "+", scrolls the new tab into view — and the "+" no longer scrolls out of reach when the strip is full, since it stands beside the tabs rather than at the end of them

### Fixed

- A middle-click on a file tab closes it again once the strip has a scrollbar. The browser had been reading that press as the start of its own auto-scroll and swallowing the gesture, which is why it worked until the tabs filled up

## [0.3.16] - 2026-10-16

Rows and columns can be hidden while you read, and the decoration a cell's own font carries — italic, underline, strikethrough — is drawn.

### Added

- Hide and Unhide on a row or column heading, from its right-click menu, with the accelerator letters Excel names them by. What is hidden this way is the reader's, not the file's: nothing is written back, and a sheet comes up as it was left for as long as it is open
- Rows and columns the file itself hides are respected. A hidden row now takes no room and its number is not drawn, so the numbering jumps exactly where Excel's does
- Italic, underline and strikethrough are drawn. 44 of the 60 workbooks sampled use italic and 42 use underline, and every one of those cells drew plain; the rules' geometry differs by font family and was read off Excel's own PDF
- A cell set to shrink its text to fit does. Excel scales the whole font down until the line fits the column, and a cell holding more than one line keeps its size, as the standard says
- Rich text draws in each run's own font. A word inside a sentence can be bold, italic, underlined, struck, another family or a size larger, and what a run says overrides the cell's own decoration — which is what leaves only the link inside a breadcrumb underlined

### Improved

- The right-click menu offers Fill colour as a menu item opening onto the theme palette, the standard one, the recent swatches and No Fill, the way Excel's does, rather than listing the swatches bare

### Fixed

- Right-clicking a heading acts on the selection. The command took the single heading the pointer happened to be over, and the sheet moves under the pointer as the view is kept in step with the selection, so a right-click aimed at J could land on N and hide one column of three
- Hide takes every row or column the selection spans — including a block dragged across cells, which used to collapse to the one row the menu was opened on — and Unhide reaches the hidden run beside the selection, since a hidden row has no heading to point at. The heading a menu was opened on is forgotten when it closes, instead of outliving it into the next command
- A merged region lights up only from its anchor cell, and a whole column's band is drawn across the merges it crosses rather than leaving them alone. Excel treats a merge as one cell, living where its text is
- A Center-Across run ends at the next cell holding a value. Two labels side by side were centred as one across the pair
- Hiding a row survives a sheet switch
- A line break in a value that does not wrap is drawn as nothing, as Excel draws it. A cover row written as two lines lost its second line

## [0.3.15] - 2026-10-15

The things a workbook keeps out of sight are marked: the rows a rule colours,
the sheets hiding past either end of the tab strip, and the rows a filter has
taken away.

### Added

- Conditional formats are drawn. This is what a review checklist greys its
  out-of-scope rows with: each rule's formula is worked out per cell, and the
  fill, font colour and border it asks for are painted where it matches. A rule
  written as a formula is what this reads — a colour scale, a data bar, or a
  rule reaching into another sheet is left alone
- The tab strip carries Excel's "…" at either end when sheets are hiding past
  it. Pressing one makes the sheet just beyond what is on screen current and
  scrolls its tab in, so pressing it again walks the selection along one sheet
  at a time; each end keeps its slot when nothing is hidden that way, so the row
  never shifts. Ctrl+click an arrow to jump straight to the first or last sheet,
  and the sheet list a right-click opens now appears at the pointer
- The status bar says how many sheets the workbook holds, beside its name
- The row numbers of the rows a filter is working on are blue, as Excel draws
  them. The grid was already skipping the numbers of the rows it hides away; the
  colour was the rest of that mark

### Improved

- What a filter is hiding is said in the status bar rather than in a band across
  the top of the grid, which spent a row of the sheet on saying it. A filter is
  cleared from its column's own funnel menu, a column at a time

### Fixed

- The sheet's scrollbar splitter stops where the tab row can take no more. Past
  that point the handle used to stand still under a pointer that was still
  moving, and the first stretch of a drag back did nothing at all; on a wide
  window, the first touch of it also snapped the bar shorter
- The search list keeps the arrow keys when the walk crosses into another sheet.
  A sheet switch hands the keyboard to the sheet — that is what makes the
  shortcuts act on it — and the list was not excepted from it, so the next press
  moved the cell instead of the next hit. It only happened once the focus sat on
  a row of the list, which is where a click on a hit leaves it

## [0.3.14] - 2026-10-15

A search that walks: its results answer to the arrow keys, and what has been
searched for comes back.

### Added

- The search results answer to the arrow keys. Down and Up walk them a row at a
  time with the sheet following, so reading down the list is looking at the
  cells it found; the row the walk is on is marked, and the list keeps it in
  sight as the walk goes past. The ends hold rather than wrapping. The list's
  scrollbar is drawn always rather than only under the pointer, since it is how
  far down the walk has got
- The search field is a combo. The button in its right edge opens what this
  session has searched for, newest first, as wide as the field and hanging
  under it; picking one puts it back in the box and searches again. They are
  the app's own memory and nothing more — nothing is written anywhere, so
  closing the app is the whole of the forgetting

### Improved

- Escape closes the search panel only when it is pressed inside the panel
  itself. It used to close it from wherever the key was pressed, so an Escape
  meant to take the marquee off a copy, to put away a cell's text or the name
  box, or to dismiss a dialog took the search away with it. Ctrl+F still closes
  it from anywhere

### Fixed

- The box a double-click or F2 opens over a cell follows a filter that hid the
  rows above it. It used to stay where its row would have been with nothing
  hidden, which on a filtered sheet left it over whatever cell had moved up into
  that space

## [0.3.13] - 2026-10-15

Excel's own colour palette, a colour for a worksheet tab, F2, a workbook the
start page can pin — and a selection whose four sides are the same weight.

### Added

- The colour palette is the one Excel draws: the theme's colours over the forty
  tints and shades that hang from them, the standard band under those, the
  colours just used under that, and No Fill at the foot. The ten slots of
  recently used colours are shared by everything that opens the palette, so a
  colour picked for a cell is one a worksheet tab may want next
- A worksheet tab takes a colour from its own right-click menu — Excel's sheet
  menu, as far as this app has anything to put in it. The colour is the app's
  own: there is nowhere to write it back to, so it lasts as long as the workbook
  is open, and taking it off hands the tab back to the colour the file gives it
  rather than to no colour at all
- F2 opens the active cell's text, which is the keyboard's way to what a
  double-click on a cell does. An empty cell has nothing to offer, so it is left
  alone
- The start page can fix a workbook to the top of the list, above the history. A
  fixed workbook does not age out with the fifty the history keeps, and clearing
  the history leaves it fixed

### Fixed

- A selection's four sides are the same weight now. The rule was drawn a pixel
  inside the cell, while the cell's own border is drawn on its boundary — so on
  a sheet with borders the right and bottom sides showed the green line *and*
  the border beside it, and read as twice as thick as the left and top. The rule
  is heavier than it was, as Excel's is, and the cell wears it while it is being
  edited as well as while it is selected
- Opening a cell no longer loses its colour. The field put over it painted the
  file's own fill and nothing else, so a cell you had marked lost its colour the
  moment you opened it, and a cell outside the print area lost its grey

## [0.3.12] - 2026-10-14

Shift+click and Ctrl+click select the way Excel's do, the window keeps answering
while a workbook is read, and a selected row takes its colour the whole way
across.

### Added

- Shift+clicking a cell selects the block between it and the cell you started
  from, and that first cell stays the one outlined: a further Shift+click grows
  or shrinks the block around it rather than around the cell just clicked. A
  click back inside the block makes it smaller instead of taking a corner off,
  and Shift+clicking while a whole row or column is selected keeps the whole
  line. The arrow keys leave from the outlined cell, as they do in Excel
- Ctrl+click adds what it lands on to the selection, and takes it off again when
  it is already in it — a cell, or a whole row or column clicked on its header.
  Ctrl+dragging pulls a block out the same way. The name box lists the areas, the
  status bar counts them, Ctrl+C copies all of them, and the paint bucket shades
  all of them

### Improved

- The window keeps answering while a workbook is opened or a sheet is read:
  every command runs off the UI thread now, rather than holding the message loop
  for as long as it runs
- A sheet's pictures arrive in one call instead of one call each, so a document
  with fifty screenshots no longer opens its archive over and over

### Fixed

- Selecting a row on a sheet scrolled past its last column left gaps in the
  row's colour: the columns whose place in the blank area lined up with a merged
  region further down the sheet were skipped. The row takes the colour all the
  way across now, and on a sheet with gridlines on, those same columns are ruled
  again
- An update rewrites the `update-manifest.json` beside the executable now. It
  used to stay at the version the folder was first unpacked from — a copy
  updated through a dozen releases still said 0.1.0 — and it names the build
  actually beside it now

## [0.3.11] - 2026-10-14

A workbook opened while another whose sheets share its names is already open
keeps its own cells now.

### Fixed

- Opening a second workbook whose sheets are named the same as the first's left
  the first one's data in it: the two shared one store of cells, and the second
  workbook's were never read. A label sitting at one column in one file turned
  up under another file's column header, and stayed there until the app was
  restarted. Each workbook's cells are its own now.

## [0.3.10] - 2026-10-14

The window with nothing open is a start page, a file tab's place in the strip
is yours to set, and shading lands on all of a selection.

### Added

- The empty window is a start page now: a row to open a workbook, the workbooks
  opened lately down the left — the File menu's own list, so a workbook taken
  out there goes from here too, each saying which folder it is in — and every key
  the app answers to down the right
- The rest of that list is one More row under the newest few
- A file tab can be dragged to another place in the strip. The order is yours,
  and the workbook showing stays the one showing
- A sheet tab's name opens over it when it is double-clicked: selectable and
  copyable, Ctrl+A taking the whole name, and nothing able to change it.
  Escape, Enter, or a press anywhere else puts it back

### Fixed

- The File menu was as wide as the word on it, so every workbook in it came out
  with an ellipsis. It is as wide as the longest name now, and no wider than a
  quarter of the window
- Double-clicking a cell too narrow to hold its value opened a field as narrow
  as the column, which held nothing. The field is the box the value is drawn in:
  as wide as the value needs, never past the neighbour that stops the spill
- Shading a row or a block on a sheet narrower than the window left the
  left-hand columns of the selection unshaded, and whether they took the colour
  depended on where the sheet was scrolled to. The mark now lands on every cell
  it was put on, and stays there
- The paint bucket stopped at the last column holding a value, so a column the
  file keeps alive with a border, and a merged region reaching past it, were left
  out. A fill now reaches as far as the sheet is drawn

## [0.3.9] - 2026-10-14

Cells can be shaded while you read, the sheet opens in whichever of Excel's two
views suits the workbook, and the sizes you change stay changed.

### Added

- Cells can be shaded in a colour, for marking a sheet up as you go: the paint
  bucket fills the selection with the colour it is holding, the caret beside it
  opens the palette, and a right-click on the grid opens the same one. The
  colours are the app's own — nothing is written back to the file
- Two view buttons at the right of the status bar, as Excel has them: the
  ordinary view, and Page Break Preview with the print area shaded around and
  boxed. A workbook that sets no print area opens in the ordinary view, and each
  sheet can be switched on its own
- A file tab's right-click copies the workbook's name, or its full path

### Improved

- The status bar no longer repeats the cell's address, its type and its value —
  the name box and the formula bar sit directly above the sheet and say the
  same. How far a selection reaches now reads beside the zoom
- Column widths and row heights you drag or fit are kept while the workbook is
  open, so leaving a sheet — for another sheet, or for another workbook — and
  coming back no longer puts the file's own sizes back
- Text too long for its cell spills past the last column the file uses, as Excel
  lets it, rather than being cut off at the edge of the data

### Fixed

- Ctrl+A while a cell's text is open over it selects that text, every time it is
  pressed, instead of walking out to the region and then to the whole sheet

## [0.3.8] - 2026-10-13

A cell's text can be read and copied without leaving the grid, rows are fitted
to what is in them, and a filter lands where the workbook says it does.

### Added

- A cell's text opens over the cell when it is double-clicked — selectable and
  copyable with the mouse, its coloured runs and all, and not changeable
- A row's height is fitted to what is in it by double-clicking the bottom border
  of its header
- The File menu's list of recent workbooks can be emptied of an entry, or
  cleared whole

### Improved

- A Shift-click on a row or column header extends the run, the way it does on a
  cell
- Dragging a row height with several rows selected sets them all, and a
  double-click on the border fits each of them to its own content
- A sheet keeps the filter it was left with when you leave it and come back,
  as Excel does

### Fixed

- The filter button sits on the row the workbook's own filter names. A sheet
  with saved custom views was read from the last of those instead: a WBS whose
  filter is `M1:FC66` was taken as some view's `A7:EY121`, seven rows lower
- A filter hides the rows inside its range and nothing else. What sits under a
  table used to vanish with it
- The filter banner's row count and its Clear button are translated like the
  rest of the interface

## [0.3.7] - 2026-10-13

Recent workbooks are one click away, and the sheet tabs are painted the way
Excel paints them.

### Added

- A File menu, where Open used to be: open a workbook, open one of the fifty
  opened lately — ten listed in the menu and the rest under More — or close the
  one showing
- A tab is closed by a middle click, the gesture browsers put on one
- A recent whose file has gone is greyed, with the reason in its tooltip.
  Clicking it tries again, and opening it puts the entry right

### Improved

- A sheet the workbook gives a tab colour is filled with it, name and all, in
  white type or black as the colour needs. The current sheet is not filled: it
  keeps the raised tab, washed in that colour with the rule along its foot
- The interface's green is the one Excel draws today, `#107C41`, where it had
  Excel 2013's `#217346`
- The toolbar gave up its Search button. Ctrl+F opens the search panel, and the
  button was the second way to do one thing
- A coloured sheet tab lights up when the pointer is over it, as the plain ones
  already did

### Fixed

- The More submenu stays inside the window: once it is taller than the space it
  has, it scrolls, rather than drawing its last entries out of reach

## [0.3.6] - 2026-10-13

The sheets can be picked from the bottom too, the way Excel keeps them.

### Added

- A strip of sheet tabs along the foot of the sheet area, beside the list on the
  left. Either can be shown or hidden on its own, and the choice is remembered
- Each tab carries the colour the workbook gives it, so the sheets are told
  apart the way they are in Excel
- The name box takes a reference: type `C16`, or `Sheet1!C16` to go to another
  sheet with it, and press Enter

### Improved

- While the tabs are showing, the sheet's horizontal scrollbar sits at the end
  of the strip, and the handle beside it makes it longer or shorter
- The formula bar's scrollbar is the same one as everywhere else in the app
- The toolbar has given up its Go To button — the name box and Ctrl+G were two
  ways to do the same thing already

### Fixed

- Scrolling no longer slows down as it passes the last of the data. The view was
  asking for the empty cells beyond it a block at a time and waiting on the
  answer

## [0.3.5] - 2026-10-13

The empty part of a sheet behaves itself.

### Improved

- Past the end of the data the scrollbar moves the way Excel's does: the range
  stands still until the view reaches its end, instead of growing outwards from
  the middle

### Fixed

- A table's rules and shading no longer run on past its data into the blank part
  of the sheet
- A sheet's drawings stay inside the sheet instead of being drawn again out to
  the right of the data
- The first frame after switching sheets is right first time, without having to
  scroll it to wake it up

## [0.3.4] - 2026-10-12

A sheet no longer ends where its data does.

### Added

- A sheet can be scrolled past its last cell, into the empty part of it, as
  Excel's can be. The blank area is made as the view reaches its end and given
  back as the view leaves it

### Fixed

- A copied merged cell is outlined around the whole merged block, not just
  around the corner the cursor sits on
- A sheet whose filter hides its last row no longer comes up grey from corner to
  corner
- The update prompt shows a release's notes as they are written, without the
  markers that group them

## [0.3.3] - 2026-10-12

Ctrl and the arrows now move the cursor the way Excel's do, and a copy says so.

### Added

- Ctrl+arrow jumps to the edge of the data — the far end of a run of filled
  cells, or the next filled cell past a blank one — and Ctrl+Shift+arrow extends
  the selection there
- Ctrl+C draws the dashed rule Excel draws around what was copied, and it crawls
  until Escape takes it away from wherever the focus happens to be

### Fixed

- Switching away from a sheet and back no longer moves the view: each sheet
  keeps the position it was left at
- A copy no longer leaves the keyboard pointing at nothing

## [0.3.2] - 2026-10-12

Two things that went wrong on a long sheet.

### Fixed

- Borders no longer stop partway down a heavily formatted sheet. The parser
  gave up on reading styles past a certain point, and what it dropped read as
  missing borders to the right of the table
- The row numbers stay out of the corner where the row and column headers meet
  while the sheet is scrolled

## [0.3.1] - 2026-10-12

Updating replaces more of the folder now. It used to swap only `SheetLens.exe`,
leaving the README and the changelog beside it describing whatever version
happened to be unpacked there first.

### Improved

- An update rewrites the README, the changelog and the licence in the folder too

### Fixed

- The README now describes the grid as it is — drawn on a canvas since 0.3.0,
  not rendered in CSS

## [0.3.0] - 2026-10-12

The grid is painted onto a canvas instead of being built out of an element per
cell, which is what used to make a long sheet slow to scroll and slow to click
in. The rest of this release is the keyboard catching up with it.

### Added

- Shift with the arrow keys, Page Up/Down or Home/End extends the selection, and
  Ctrl+A selects the data around the cursor — the whole sheet on a second press
- Dragging along the row or column headers selects a run of them, `3:7` or `C:E`
- Ctrl+Home and Ctrl+End jump to the two corners, Ctrl+Page Up/Down change sheet

### Improved

- Large sheets are responsive again: scrolling, selecting and clicking no longer
  stall the way they did
- A new icon

### Fixed

- Ctrl+Shift with a letter no longer fires the plain shortcut under it —
  Ctrl+Shift+G was opening the ordinary Go To

## [0.2.0] - 2026-10-11

The interface now comes in English, Chinese and Japanese, and opens in your
system language. The button at the right of the toolbar switches it, and what
you pick is remembered.

### Added

- Simplified Chinese and Japanese throughout
- A language switcher in the toolbar
- The notes shown when an update is available are in that language too

## [0.1.1] - 2026-10-11

An update now says what the release it is offering contains.

### Improved

- The prompt that offers an update shows that release's own notes, in place of
  the placeholder every one of them had carried
- The update dialog shows how far the download has got, instead of a line that
  only said "downloading"

## [0.1.0] - 2026-10-11

First release: a portable Windows build of SheetLens that can update itself in
place. Unpack it anywhere and run `SheetLens.exe`; it keeps itself current from
the releases on this repository.

### Added

- Opens `.xlsx`, `.xlsm`, `.xls`, `.xlsb` and `.ods`
- Sheets, search, go-to, column filters and a formula bar
- Cell, row/column and rectangular-block selection, with Ctrl+C in Excel's own
  format
- Ctrl+wheel zoom on Excel's own steps
- Excel-faithful layout: column widths, merges, borders, fills, fonts, the print
  area, and the shapes and pictures floating above the cells
