# ExportUV — CorelDRAW VBA Macro

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
![Status](https://img.shields.io/badge/status-stable-success)
![Platform](https://img.shields.io/badge/platform-Windows-informational)
![CorelDRAW](https://img.shields.io/badge/CorelDRAW-VBA-green)
[![SPDX](https://img.shields.io/badge/SPDX-GPL--3.0--or--later-lightgrey)](https://spdx.org/licenses/GPL-3.0-or-later.html)

> Automates a small **UV prepress** workflow on the **active page** in CorelDRAW.  
> Groups artwork, applies a straight-edge **Envelope** in *Original* mode, nudges geometry, draws a transparent registration frame, and cleans up — all in one undoable step.

---

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [What It Creates](#what-it-creates)
- [Geometry \& Colors](#geometry--colors)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)
- [Limitations](#limitations)
- [Contributing](#contributing)
- [License](#license)
- [Changelog](#changelog)

---

## Features

- 🔎 Detects a **310×500 mm** rectangle (±0.05 mm) anywhere and renames it `START_FRAME`  
- 🧩 Groups all page shapes into **`MY_OBJECT`** (or reuses a single existing group)  
- ✨ Creates **`ENV_TRANSFORM`** polygon and applies an **Envelope in Original mode** (keeps straight lines)  
- 📐 Nudges everything by **+0.525 mm (X)** and **+0.2 mm (Y)**  
- 🖼️ Draws transparent registration frame **`MY_FRAME`** at precise coordinates  
- 🧹 Deletes `START_FRAME` at the end (even if it lives inside groups or PowerClips)  
- ♻️ Restores document units & wraps everything into **one command group** (single Undo)

---

## Requirements

- **CorelDRAW for Windows** with **VBA** enabled
- A document with artwork on the **active page**
- *(Optional)* A rectangle close to **310×500 mm** if you want it recognized as `START_FRAME`

> The macro temporarily switches **document units to millimeters** and restores them afterward.

---

## Installation

1. Open CorelDRAW.  
2. Go to **Tools → Macros → Macro Manager** (or press **Alt+F11**).  
3. Open or create a VBA project (e.g., `GlobalMacros.gms`).  
4. **Insert → Module**, name it e.g. `ExportUV`.  
5. Paste:
   - the **GPL-3.0 header** (see header snippet in your module), and  
   - the macro code (starting with `Option Explicit`).  
6. Save the project.  
7. *(Optional)* Bind a shortcut or toolbar button to `ExportUV`.

---

## Quick Start

1. Open your document and activate the target **page**.  
2. *(Optional)* Ensure a **310×500 mm** rectangle exists; it will be auto-renamed to `START_FRAME`.  
3. Run **ExportUV**.  
4. Verify results:  
   - All shapes grouped as **`MY_OBJECT`**  
   - Envelope transformation applied (straight edges preserved)  
   - Content nudged by **(+0.525, +0.2) mm**  
   - **`MY_FRAME`** is present (transparent)  
   - **`START_FRAME`** removed (if it existed)

> Everything is a **single Undo** operation thanks to `BeginCommandGroup`/`EndCommandGroup`.

---

## What It Creates

- **`MY_OBJECT`** — group containing all shapes on the active page (or the one existing group)  
- **`ENV_TRANSFORM`** — temporary polygon used for the Envelope (deleted after use)  
- **`MY_FRAME`** — transparent rectangular frame for downstream layout/fixtures

---

## Geometry & Colors

**Envelope Polygon Points (mm):**

| Point | X (mm) | Y (mm) |
|:-----:|-------:|-------:|
| 1 | -47.25  | 398.496 |
| 2 | 261.1   | 399.65  |
| 3 | 258.393 | -100.644 |
| 4 | -50.05  | -102.25 |

- Outline (for visibility before deletion): **CMYK(100, 100, 0, 0)**  
- The polygon is **deleted** after the Envelope is applied.

**Nudge:** `+0.525 mm (X)` and `+0.2 mm (Y)` (applies to all shapes on the active page).

**Registration Frame (`MY_FRAME`) Corners (mm):**

| Corner | X (mm) | Y (mm) |
|:------:|-------:|-------:|
| Top-left | -50   | 398.5  |
| Bottom-right | 260 | -101.5 |

- Fill: **No Fill**  
- Outline: **None** (fully transparent)

---

## Configuration

Adjust these constants/lines in the code as needed:

```vb
' Tolerance for size match
Const TOL_MM As Double = 0.05

' START_FRAME target size (mm)
Const TARGET_W_MM As Double = 310#
Const TARGET_H_MM As Double = 500#

' Nudge
srAll.Move 0.525, 0.2

' Registration frame
Set sFrame = lyr.CreateRectangle(-50#, 398.5, 260#, -101.5)

' Envelope polygon points
Set sEnv = CreateClosedStraightPolygon(lyr, Array( _
    Array(-47.25, 398.496), _
    Array(261.1, 399.65), _
    Array(258.393, -100.644), _
    Array(-50.05, -102.25) _
))
```

---

## Troubleshooting

- **“No shapes on the active page.”** — The macro fails fast if empty; add artwork or switch pages.  
- **Locked/hidden layers** — Unlock/show layers; the macro groups **all** shapes on the active page.  
- **Single existing group** — Reused as `MY_OBJECT`, not regrouped.  
- **`START_FRAME` not found** — That’s fine; the final delete step simply does nothing.  
- **Export errors** — This macro does **not** call `ExportBitmap`. If you saw that error, it’s from another macro.

---

## Limitations

- Operates on the **active page only**  
- Groups **all** shapes on that page (customize if you need per-layer filtering)  
- Assumes standard CorelDRAW VBA constants (numeric fallbacks are used where helpful)

---

## Contributing

Issues and pull requests are welcome. If you submit changes, please describe **why** and include **before/after** notes or screenshots when relevant.

---

## License

**GPL-3.0-or-later** • See the included `LICENSE` file or read at [gnu.org/licenses](https://www.gnu.org/licenses/).  
SPDX identifier: `GPL-3.0-or-later`

---

## Changelog

### 1.0.0 — 2025-08-31
- Initial public release

---

**Author:** Nikola Karpić
