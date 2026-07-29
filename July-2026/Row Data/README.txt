ROW DATA — July 2026
====================

Drop this month's photos into the folders below. That is all that's needed —
the sheet's "Drive Link" column can be left completely empty for images. These
folder names ARE the wiring; nothing has to point at them.

Order within a folder is filename order, so prefix with 01, 02, ... when it
matters.

A folder you leave empty falls back to whatever the sheet links (if anything),
which for now still means last month's GitHub images — so an empty folder can
quietly show the wrong month's photos. The run prints where every section got
its images from; check those lines read "from Row Data" before sending.

IMPORTANT: this folder is found by name. The sheet's month/year must read
July / 2026 for "July-2026" to be picked up. Change the sheet first, or the
scripts will look for a folder named after whatever month the sheet says.


FILENAME MATTERS IN THESE THREE
-------------------------------
The text shown in the newsletter is read from the filename — there is no sheet
row for these people.

Separator: either "--" (two dashes) or " - " (hyphen WITH a space either
side). Both work, so all of these are fine:

    Arpi Shah - Sales Automation Analyst.jpeg
    Arpi Shah--Sales Automation Analyst.jpeg

The spaces around a single hyphen are what matters — they are how a
hyphenated name like "Anne-Marie" is kept intact instead of being split.
A file with no separator at all is skipped, and the run prints a WARNING
naming the file so it never disappears silently.

awards/          01 - Award Title - Person Name - Designation.jpg
                 Leading number = display order, and is not shown.
                 Repeat the same Award Title across files to group several
                 people under one award, e.g.
                   04 - Team Spirit Award - Pooja Dahale - Sr. Odoo Consultant.png
                   04 - Team Spirit Award - Uttam Jain - Lead Odoo Consultant.png

new_joinee/      Person Name - Designation.png

events/event1/   Event photo files (any names). event1 / event2 / event3 map
events/event2/   to the three HR Insider events; their TITLES come from the
events/event3/   hr_event1 / hr_event2 / hr_event3 rows in the sheet.


ANY FILENAME IS FINE IN THESE
-----------------------------
new_customer/    Customer logos
certification/   Certificate images
announcement/    Upcoming event banners
excellence/      Client logos for the Acknowledging Excellence quotes
delivery/        Client logo for Delivery Insights
marketing/blogs/ Blog thumbnails — blog1.jpg, blog2.jpg (order matters here;
                 titles and links come from the blog1 / blog2 sheet rows)


ONE IMAGE ONLY — the first file in the folder is used
-----------------------------------------------------
ceo/             CEO photo
campaign_banner/ Campaign / Whistle Blower banner
hr_wellness/     Wellness & HR Corner banner


A section with no images and no sheet text is skipped entirely — it will not
appear in the preview or the final email.

Accepted formats: .jpg .jpeg .png .gif .webp .svg
Avoid very large files: a 17 MB photo makes the flattened section PNG huge and
slow to load in the email.
