FileConverter
=============
Right-click a file -> Convert to -> PDF / JPEG image / Word document.
The new file appears next to the original. Nothing is ever overwritten
("report.pdf" becomes "report (1).pdf" if one already exists).

INSTALL:   double-click install.bat   (no admin needed; installs Python if missing)
UNINSTALL: double-click uninstall.bat

Windows 11: the menu lives under "Show more options" (or Shift + right-click).
Multi-select works: select 20 files, right-click once.

What converts to what
---------------------
Images (png, bmp, gif, tiff, webp, heic, jpg)  -> PDF, JPEG, Word
PDF                                            -> JPEG, Word
Word (docx, doc, rtf, odt)                     -> PDF, JPEG (+ doc/rtf/odt -> docx)
PowerPoint (pptx, ppt, odp)                    -> PDF, JPEG
Excel (xlsx, xls, ods, csv)                    -> PDF, JPEG
Text (txt, md)                                 -> PDF, JPEG, Word

Multi-page sources -> JPEG produce a folder "name (JPEG)" with one image per page.
Office formats need Microsoft Office or LibreOffice installed.
PDF -> Word reconstructs layout; complex PDFs (scans, heavy tables) come out imperfect.

Problems? See %LOCALAPPDATA%\FileConverter\log.txt
Do not move this folder after installing - or run install.bat again if you do.
