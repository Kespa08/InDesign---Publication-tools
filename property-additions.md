# Selection Inspector — Property Additions

Use this file to request new properties for the script.
Fill in one or more entries below, then ask Claude to implement them.

---

## How to fill out an entry

| Field | What to write |
|---|---|
| **Mode** | Which object type: `Text` / `Cell` / `Table` / `Frame` / `All` |
| **Property** | The InDesign UI label or a plain-language description |
| **Type** | `numeric` — a measurable quantity with a tolerance field (e.g. stroke weight, point size) |
| | `style-name` — a named style applied to the object (e.g. paragraph style, object style) |
| | `enum` — one of a fixed named set (e.g. baseline offset mode: Ascent / Cap Height / Fixed) |
| | `boolean` — on or off (e.g. overprint fill) |
| **Units** | If numeric: the label to display in the UI (`pt`, `mm`, `%`, etc.) |
| **Notes** | Optional. Typical values you have seen, any edge cases, preferred match behaviour. |

Every property added to the script appears in both Stage 1 (matching) and Stage 4 (adjusting).
There is no "display only" option — the vocabulary is the same in both stages.

---

## Pending entries

<!-- Add new entries here. Claude will implement them and move them to Implemented. -->


---

## Implemented entries

<!-- Claude moves completed entries here after implementation, with the date. -->
