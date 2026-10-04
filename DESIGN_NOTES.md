# Design study → design families

The pictures in `PPT_Designs/` were studied for their **building blocks** — shapes, type,
placement, rhythm — and those were rebuilt as original slide compositions in `make_ppt.py`.
Nothing is traced or copied; each family is a fresh design drawn with PowerPoint shapes.

## What the 17 references are made of

| # | Reference | Shapes | Type | Placement & rhythm |
|---|---|---|---|---|
| 1 | Concentric arc infographic | stacked quarter-ring arcs, numbered ends (01–04), small icon circles with leader lines to % callouts | condensed uppercase labels, big light numerals | hub bottom-left, arcs fan out; data callouts float on the left with thin connector lines |
| 2 | Soft UI "Customize Infographic" | white rounded cards with soft shadows on light grey, ring gauges, thin bar charts, a raised circle with the page number | geometric sans, heavy title with a lighter second word | three tiny coloured dots top-left, title under them, page number circle top-right; generous margins |
| 3 | Hexagon cluster | white hexagons with a thick coloured arc on the inner edge, dashed curved arrows, centre map | condensed caps title + light lowercase subtitle, coloured dots under it | three hexagons orbit a centre; title block to the left |
| 4 | Editorial law-firm deck | black/cream panels with diagonal cuts and curved masks, photo halves, small hexagon portraits | serif display headlines, small caps kickers | strong left/right split, lots of white space, thin gold rules |
| 5 | Soft UI "Technology" | same soft cards; drop/teardrop shapes with numbers, donut charts | geometric sans, title split bold/light | 2-column card grid, coloured accents on small elements only |
| 6 | Timeline collection | curving road, arcs of years, staircases, arrow bands, milestone dots | bold sans headings | horizontal flow left→right, alternating above/below labels |
| 7 | Modern S-curve infographic | full-width banners alternating left/right, joined by an S-shaped ring path with icon circles | condensed caps "INFOGRAPHIC 01" headings | zig-zag reading order; colour steps from teal to navy |
| 8 | Segmented ring | pie cut into wedges with an outer coloured rim, centre hub circle, big numbers outside | condensed caps, big coloured numerals | numbers sit outside the ring, each with a short label |
| 9 | Corporate photo deck | photo bands, giant section numerals (1, 2, 3), hexagon icon groups, kicker bar "MORE THAN TEMPLATE" | condensed caps headings, big thin numerals | dark photo band on top, content on light paper below |
| 10 | Ribbon steps | tall coloured ribbons with folded ends, "01 STEP" stacked numerals | bold condensed numerals | vertical ribbon spine, text alternating left/right |
| 11 | Travel magazine | large photos with organic curved masks, circular teal icon badges, small rounded image cards with captions | serif two-line headlines, small caps kickers | text column left, image right; warm off-white paper |
| 12 | Glassy web UI | frosted cards, gradient bar charts, donuts | modern sans | dashboard grid, pastel gradients |
| 13 | Soft UI charts | same soft language: blue gradient rings, half-donuts, area charts, big % numbers | geometric sans | numbers dominate, captions tiny |
| 14 | Black & white business | black panels, outlined title frame, circles in a 2×2 business model, SWOT ring | bold sans caps title in a thin frame | high contrast; one steel-blue accent |
| 15 | Teal corporate | droplet hub with ring of segments, map pins on a curve, corner brackets | high-contrast serif display + sans body | big image tops, corner bracket ornaments |
| 16 | Diamond steps | rounded diamonds with a coloured triangle cap ("STEP 01"), connector lines from a big arrow | bold condensed caps | diagonal zig-zag of diamonds |
| 17 | Hexagonal swirl | six interlocking curved blades forming a ring, titles A–F around it | light sans, big numerals | radial, six-way symmetry; one hot accent on grey |

## Recurring ideas worth building

- **Numbers as graphics**: oversized step/section numerals (1, 9, 10, 15, 16).
- **Radial structures**: rings, arcs, segments around a hub (1, 3, 8, 13, 15, 17).
- **Geometric containers**: hexagons, diamonds, ribbons, circles instead of plain boxes (3, 9, 10, 16, 7).
- **Soft depth**: white cards with large blurred shadows on a light grey canvas (2, 5, 12, 13).
- **Split compositions**: half dark panel / half light, often with a diagonal or curved edge (4, 11, 14).
- **Tiny decorative systems**: three coloured dots, corner brackets, kicker bars, page-number circles.
- **Type pairing**: condensed caps for infographics; serif headlines for editorial/travel; geometric sans for UI.

## The families (original compositions)

| Family | Built from ideas in | Signature |
|---|---|---|
| **Soft Dashboard** | 2, 5, 12, 13 | light grey canvas, raised white cards, three-dot + page-number header, ring-gauge stats, Century Gothic titles |
| **Ribbon Steps** | 10, 16, 7, 1 | crisp white, folded colour ribbons with "01 STEP" numerals, diamond cards, Bahnschrift condensed caps |
| **Hexa Hub** | 3, 8, 17, 9 | pale blue-grey, hexagon cards with coloured edges, segmented ring stats, giant numeral section slides |
| **Editorial Mono** | 4, 14, 15 | black/white with one steel accent, diagonal split title, Georgia headlines, thin rules, framed covers |
| **Wanderlust** | 11 | warm paper, serif two-line headlines, organic photo masks, teal circle badges, pill tags |
| **Timeline Flow** | 6, 15, 9 | chevron arrow steps, milestone dots on a line, kicker bar headers, bold numbers |

Fonts are all installed with Windows (Century Gothic, Bahnschrift, Georgia, Segoe UI), so decks
render the same offline.
