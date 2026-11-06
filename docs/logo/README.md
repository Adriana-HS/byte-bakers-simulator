# Byte Bakers — logo and visual identity

Owner: **Misael** · Task [#24](https://github.com/Adriana-HS/byte-bakers-simulator/issues/24) (M1) · Rubric **G7 / G6**

## Concept

The computer is the kitchen and data are the ingredients on the buses.  
The mark is a **chef hat whose band is an 8-pin chip** (one pin per bit of a byte), as suggested in the team plan (section 10).

## Files

| File | Use |
|---|---|
| `logo.svg` | Icon (vector) |
| `logo.png` | Icon 512×512 (app / docs) |
| `logo-32.png` | Icon 32×32 (readability check) |
| `logo-transparent.png` | Icon on transparent background (sheet overlay) |
| `logo-horizontal.svg` | Logo + wordmark (vector) |
| `logo-horizontal.png` | Logo + wordmark 1024×256 |

## Palette (matches `modConfig` / plan §3.7 and §10)

| Role | Hex (UI) | VBA `Const` (BGR) | Meaning |
|---|---|---|---|
| Background | `#1F2937` | `CLR_BG` `&H37291F&` | Dark blue sheet |
| Accent / active | `#D85A30` | `CLR_ACTIVE` `&H305AD8&` | Coral — active driver / chip band |
| Text / cream | `#F1EFE8` | `CLR_TEXT` `&HE8EFF1&` | Off-white text and hat |
| Load / amber | `#EF9F27` | `CLR_LOAD` `&H279FEF&` | Chip pins |
| Idle / High-Z | `#B4B2A9` | `CLR_HIGHZ` `&HA9B2B4&` | Gray underline wire |

Do **not** invent extra brand colors outside this table for the simulator UI.

## Typography

- Headings / wordmark: Montserrat, Poppins or **Segoe UI** (Windows).
- Hex values, LCD and bit labels: **Consolas** / Cascadia Mono.

## Where the logo goes

| Surface | Status |
|---|---|
| `README.md` / `README_EN.md` | Placed in this PR |
| Simulator sheet | Consumed by **M3** (`BuildDashboard`); asset ready (`logo-transparent.png`) |
| Defense slides | Consumed by **M11** (#34); same assets |

## Defense outfit

Kitchen apron with this logo over a formal shirt (rubric I3 team outfit).
