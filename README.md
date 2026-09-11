# Nuke All Appearances

A SOLIDWORKS macro that strips every appearance of every kind, from every display
state, every configuration, and every part inside an assembly.

**Tested on SOLIDWORKS 2022, 2024 and 2025.** Other versions are likely to
work; they are simply untested.

> **Want all eleven macros at once?** This one is part of the [MacroShelf
> Collection](https://github.com/james-debono/macroshelf-collection-sw-macro-library) — a single
> download, already arranged as a library for the [MacroShelf](https://github.com/james-debono/macroshelf-sw-addin)
> add-in, so every macro lands on a SOLIDWORKS toolbar tab with its icon and
> hover text already set up.

## What it does

Everything the other removal macros stop short of. Face, feature, body, part and
component appearances all go — in **every display state**, in **every
configuration**, and in an assembly, **inside every part file it references**.

That last point is the one to understand before you run it. The other appearance
macros deliberately stop at the assembly boundary, because reaching further means
changing documents you never opened. This macro reaches further on purpose. **A
part used in three assemblies is changed for all three.**

### Nothing is saved

The parts are modified in memory and left **unsaved**. SOLIDWORKS asks about them
when you close, so:

- **Save** — the change becomes permanent. There is no undo.
- **Close without saving** — everything goes back the way it was.

That is the safety net, and it is the only one. Appearance changes leave no undo
record, so once a file is saved the colours are gone for good.

### Suppressed and lightweight components

Both are resolved first, because an appearance cannot be cleared from a component
that is not loaded. They stay resolved afterwards — which rides on the same
bargain as everything else, so closing without saving puts the suppression states
back too.

### Which removal macro do I want?

| | Clears | Display states | Touches part files |
|---|---|---|---|
| [Remove Body and Component Appearances](https://github.com/james-debono/remove-body-and-component-appearances-sw-macro) | body and component level only | active only | no |
| [Remove All Appearances](https://github.com/james-debono/remove-all-appearances-sw-macro) | everything in the open document | active only | no |
| **Nuke All Appearances** | **everything, everywhere** | **all of them** | **yes** |

Reach for one of the first two unless you genuinely mean the third.

## Install

**The macro on its own:** download `Nuke-All-Appearances.swp` from the
[latest release](../../releases/latest), then run it with **Tools > Macro > Run**,
or add it to a toolbar with **Tools > Customize > Commands > Macro**.

**With [MacroShelf](https://github.com/james-debono/macroshelf-sw-addin):** get the
[MacroShelf Collection](https://github.com/james-debono/macroshelf-collection-sw-macro-library/releases/latest),
which packages this macro with its icon and hover text alongside every other macro
in the set. Point MacroShelf at the unzipped folder and it appears as a button.

## Using it

Open a part or assembly and run the macro. The prompt spells out exactly what is
about to happen — read it, because the assembly case is broader than most macros.
Confirm, and it reports how many appearances it cleared and how many documents it
left modified.

Then decide whether to save.

## Limitations

- **No undo.** Appearance changes leave no undo record. Closing without saving is
  the only way back.
- **Read-only and PDM-controlled parts** are still modified in memory, but you
  will not be able to save them. Nothing warns you until you try.
- **Large assemblies take time.** Every referenced part is opened and swept. The
  per-face sweep is the expensive part and can be turned off in the source — see
  [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).
- **Pattern and mirror features can put colour back.** If "Propagate visual
  properties" is on, a rebuild re-stamps the seed body's colour onto the derived
  faces. That is the feature working as designed, not the macro failing — see the
  development notes.

## Related macros

- [Remove All Appearances](https://github.com/james-debono/remove-all-appearances-sw-macro)
  — everything in the open document, active display state only
- [Remove Body and Component Appearances](https://github.com/james-debono/remove-body-and-component-appearances-sw-macro)
  — the targeted removal, which keeps face and feature colours
- [Apply Unique Colours](https://github.com/james-debono/apply-unique-colours-sw-macro)
  — colours every geometrically unique body

## Building from source

`src\Nuke-All-Appearances.vba` is the readable source. A `.swp` is a binary VBA
project, so it can only be produced from inside SOLIDWORKS — there is no build
step. Open the `.swp` via **Tools > Macro > Edit**, paste the source in, and save.

Technical detail is in [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

## Licence

MIT — see [LICENSE](LICENSE). Free to use, modify and share. The full licence text
is also carried inside the macro itself, so a `.swp` passed on by itself still
carries its licence.

Created by James Debono, with AI assistance. Everything here was tested by
hand in SOLIDWORKS — nothing that touches the API can be verified any other way.

## Trademarks

SOLIDWORKS is a registered trademark of Dassault Systèmes SolidWorks Corporation.
This project is independent: it is not affiliated with, endorsed by, or sponsored
by Dassault Systèmes, and uses only the published SOLIDWORKS API.
