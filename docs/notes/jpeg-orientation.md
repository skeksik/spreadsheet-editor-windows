# JPEG orientation in Excel

Roadmap item `spike/jpeg-orientation` (M1). Checked on 2026-11-03 with sample 09 of the Excel check (`excel-check.md`).

## The question

A JPEG can carry an EXIF orientation tag that tells a viewer to turn or flip the stored pixels to show them. When such a JPEG is a
picture in an `.xlsx`, does Excel apply the tag? If it does, a picture's shown width and height are swapped for orientations 5–8,
so `feat/picture-facts`, which reads a JPEG's size from its SOF marker, would have to swap them too. Otherwise the layout would
keep the wrong shape.

## What was checked

`09-jpeg-orientation.xlsx`, sheet **Orientation**, holds the same made-up 96 × 48 px JPEG twice, each shown 200 × 100 px. Its top
half is red, its bottom half blue, with a black square at the top left.
- B2: the plain file, with no EXIF.
- B3: the same pixels with an APP1 Exif segment whose orientation tag (0x0112) is 6, "turn a quarter turn clockwise to show". The
  bytes were checked in the saved file (`xl/media/image2.jpeg`).

Excel: **Excel for the web**, build 16.0.20626.42304, in Microsoft Edge 156. Desktop Excel is not installed on the owner's PC.

## What Excel showed

**B3 looks exactly like B2**: red top half, blue bottom half, the black square at the top left, filling the same 200 × 100 box. It
is not turned. The owner saw the same on screen.

## Conclusion

Excel for the web ignores EXIF orientation. It shows a JPEG's pixels as they are stored. So nothing changes: `feat/picture-facts`
keeps the SOF marker's width and height for every orientation, and no follow-up task is needed.

Not verified: desktop Excel (Windows or Mac). If a later check in desktop Excel shows B3 turned, the follow-up comes back:
`feat/picture-facts` swaps width and height for orientations 5–8.
