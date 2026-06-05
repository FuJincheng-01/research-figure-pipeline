# Research Figure Pipeline Checklists

## Figure Specification Checklist

- Purpose is clear in one sentence.
- Paper/source files have been inspected when available.
- Module sequence is defined.
- All exact labels and counts are extracted from authoritative source.
- Class/order labels are fixed before drawing.
- Real images needed by the figure are identified.
- Annotation type is known: BBox, mask, class label, keypoint, etc.
- Output target is Visio `.vsdx` plus exported preview. Other formats are intermediate only.

## Master Visual Review Checklist

- Overall style matches target venue/paper tone and was produced with ChatGPT image2.
- Layout reads left-to-right or top-to-bottom without ambiguity.
- One module is not accidentally over-emphasized unless intended.
- Color roles are consistent across modules.
- Arrows show only meaningful relationships.
- No irrelevant decorative objects.
- No invented values, classes, or method names.
- Generated tiny text is not trusted as final truth.
- The master is reviewed before module generation/splitting.
- The user explicitly approved the overall master before module generation started.
- If the user requested changes, a revised full master was generated and reviewed again.
- The review loop continued until the user confirmed the master was correct.

## Real Image Replacement Checklist

- User/paper-provided image is used instead of generated substitute.
- Crop keeps the relevant visual evidence.
- Annotation box matches the user's or paper's intended target.
- Box style is visible at final scale.
- Same card size and aspect ratio are preserved after replacement.
- Real image does not break the surrounding visual style.

## Module Regeneration Checklist

- Each major block is regenerated as its own ChatGPT image2 module after user approval of the full master.
- Module crops exclude cross-module arrows when those arrows will be editor-drawn.
- Module assets contain no cross-module arrows.
- Module internals use icons, large keywords, and exact numeric cards rather than long paragraphs.
- Each module preserves the approved master layout and visual semantics.
- Module regeneration improves clarity without redesigning the approved figure.
- Crop-only module assets are not used as the normal final method unless explicitly requested or needed as a fallback.
- Each module prompt records role, content, style, constraints, and exact data.
- Header/footer are separate assets when 1:1 fidelity matters.
- PNG modules are kept for visual fidelity.
- SVG wrappers or vector SVGs are kept when requested.
- Manifest records crop boxes and sizes.

## Visio COM Assembly Checklist

- Page aspect ratio matches the approved master.
- Microsoft Visio is opened or attached through PowerShell COM.
- The script uses `GetActiveObject("Visio.Application")` or creates `New-Object -ComObject Visio.Application`.
- Visio is set visible for user inspection.
- The final assembly is performed in real local Visio, not by standalone SVG/PNG/PDF composition.
- The user can see Visio open/draw when practical.
- Imported modules keep their original aspect ratio.
- Cross-module arrows are drawn in Visio and aligned.
- Final Visio layout 1:1 reproduces the approved master.
- Module and arrow positions are based on the approved master coordinates or manifest.
- Other module internals are not redrawn unnecessarily.
- Exported preview is inspected.
- VSDX package structure is checked when fidelity matters:
  - expected media count
  - expected shape count
  - expected arrows/connectors

## Readability Checklist

- Bottom notes are readable at paper scale.
- Long explanatory sentences are replaced by icons, keywords, or numeric cards.
- Count tables use large enough numbers.
- Important labels remain legible after preview export.
- Text does not overlap or touch borders.
- Footnotes are short and not essential for interpreting the figure.

## Completion Audit

Completion is proven only when:

- the user approved the overall master before module/Visio work proceeded,
- final editable/exportable file exists,
- final file is a Visio `.vsdx` assembled through local Visio COM,
- Visio COM was used as the final assembly/export mechanism,
- preview exists and has been inspected,
- preview visually matches the approved master 1:1,
- requested unchanged regions remain unchanged,
- paper data has been verified,
- real image/annotation replacements have been checked,
- asset directory and scripts are present if reproducibility was requested,
- caveats are documented rather than hidden.
