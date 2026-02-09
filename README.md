# SolidWorks 2025 Macros

VBA macros for SolidWorks 2025 that automate drawing export and assembly documentation.

---

## ExportAssemblyPDFs

A SolidWorks VBA macro that batch-exports drawing files to PDF with smart, metadata-driven filenames.

## What It Does

Instead of exporting drawings with their raw filenames, this macro pulls description text from SolidWorks custom properties to generate meaningful PDF names. For example, a drawing `10001.SLDDRW` with a description of "Mounting Bracket" becomes `10001Mounting Bracket.pdf` (or `10001 - Mounting Bracket.pdf` with a configured separator).

## Configuration

Settings are defined as constants at the top of `ExportAssemblyPDFs.vb`:

| Constant | Default | Purpose |
|---|---|---|
| `INCLUDE_TOP_LEVEL_ASSEMBLY` | `True` | Include top-level assembly in processing |
| `SKIP_SUPPRESSED` | `True` | Skip suppressed components |
| `CHECK_DRAWINGS_SUBFOLDER` | `True` | Look for a "Drawings" subfolder |
| `NAME_SEPARATOR` | `""` | Separator between base filename and description (e.g. `" - "`) |

## Functions

### `ExportDrawingToPdf(drwPath, outFolder) -> Boolean`

Main export routine. Opens a drawing silently, builds a filename from metadata, saves as PDF, then closes the drawing. Returns `True` on success.

### `GetDescriptionForDrawing(swDrw) -> String`

Resolves a description string by checking these sources in priority order:

1. Drawing's custom property `Description`
2. Drawing's custom property `SW-Description`
3. Referenced model's config-specific `Description`
4. Referenced model's custom tab `Description`
5. Referenced model's config-specific `SW-Description`
6. Referenced model's custom tab `SW-Description`

Returns empty string if no description is found anywhere.

### `GetCustomProp(doc, cfgName, propName) -> String`

Low-level helper that reads a custom property value from a SolidWorks document using the `Get4()` API. Pass `""` for `cfgName` to read from the custom tab, or a configuration name for config-specific properties.

### `CleanFileName(s) -> String`

Sanitizes a string for use as a Windows filename by:
- Replacing `\ / : |` with hyphens
- Removing `* ? " < >`
- Collapsing multiple spaces

## Requirements

- SolidWorks 2025
- SolidWorks API types used: `ModelDoc2`, `DrawingDoc`, `ModelDocExtension`, `CustomPropertyManager`, `View`

---

## ExportAssemblyPackage

Exports all fabricated components from an open assembly as PDFs + PNG screenshots and generates a self-contained HTML tree browser (`index.html`). Vendor parts (non-blank "Vendor" custom property) are automatically excluded.

### What It Does

1. Recursively traverses the active assembly tree
2. Skips suppressed components and vendor parts
3. For each fabricated part, finds its drawing and exports a PDF
4. Captures an isometric PNG screenshot of each part
5. Generates `index.html` with a collapsible tree matching the assembly hierarchy
6. Deduplicates exports — parts used multiple times are only exported once but appear at every instance in the tree

### Output

All files are written to a user-selected folder:
- `*.pdf` — drawing exports named `<PartNumber> - <Description>.pdf`
- `*.png` — isometric part screenshots
- `index.html` — collapsible tree browser with thumbnails and PDF links

### Configuration

Settings are defined as constants at the top of `ExportAssemblyPackage.vb`:

| Constant | Default | Purpose |
|---|---|---|
| `SKIP_SUPPRESSED` | `True` | Skip suppressed components |
| `NAME_SEPARATOR` | `" - "` | Separator between base filename and description |
| `SCREENSHOT_WIDTH` | `800` | Width for screenshot fallback (SaveBMP) |
| `SCREENSHOT_HEIGHT` | `600` | Height for screenshot fallback (SaveBMP) |

### Key Functions

| Function | Purpose |
|---|---|
| `Main()` | Entry point — validates assembly, prompts for folder, orchestrates export |
| `TraverseComponent()` | Recursive traversal with dedup, vendor filtering, and HTML generation |
| `FindDrawingForPart()` | Locates `.SLDDRW` files via `GetDocumentDependencies2` |
| `ExportDrawingToPdf()` | Opens drawing silently, exports PDF, closes |
| `CapturePartScreenshot()` | Sets isometric view, zoom-to-fit, saves PNG |
| `GetDescription()` | Reads Description/SW-Description from custom properties |
| `IsVendorPart()` | Returns True if Vendor property is non-blank |
| `InitHtml()` / `FinaliseHtml()` | Builds self-contained HTML with inline CSS and JS |
| `BrowseForFolder()` | Shell folder picker dialog |

### Requirements

- SolidWorks 2025
- Active document must be an assembly
- SolidWorks API types used: `ModelDoc2`, `AssemblyDoc`, `Component2`, `ModelDocExtension`, `CustomPropertyManager`
