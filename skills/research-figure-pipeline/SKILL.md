---
name: research-figure-pipeline
description: "Use when creating, reconstructing, or optimizing publication-quality research paper figures end to end for SCI papers, top-tier conference papers, theses, and technical reports across disciplines: extract paper truth, create a ChatGPT image2 master, generate module assets with ChatGPT image2, replace critical samples with real images, and assemble/export only through the user's local Microsoft Visio controlled by PowerShell COM. Applies to dataset, experiment, method, system, model, data-processing, result-analysis, and multi-module scientific figures that require modular image assets, exact paper data, Visio layout/arrows, and exported-preview validation."
---

# Research Figure Pipeline

Use this skill to build publication-grade research figures from paper content through a general workflow for SCI papers, top-tier conference papers, theses, and technical reports across disciplines:

1. extract truth from the paper,
2. create an approved master visual with ChatGPT image2,
3. wait for explicit user approval of that master visual,
4. decompose the approved master into module prompts and regenerate each module with ChatGPT image2,
5. assemble only in the user's local Microsoft Visio through PowerShell COM,
6. verify by exported preview.

The central rule: **ChatGPT image2 builds visual modules; Visio assembles relationships.** Do not hand-draw complex icons in Visio. Do not use another diagram editor as the final assembly path.

## When To Use

Use this workflow for:

- dataset construction figures
- augmentation/data pipeline figures
- model architecture figures
- method overview figures
- experiment workflow diagrams
- figures where ChatGPT image2 module assets look better than native Visio shapes
- figures requiring repeated user review and refinement

If the user references a paper, PDF, LaTeX source, or manuscript material, inspect it first. Use the paper as the authority for labels, counts, class names, order, equations, and annotation protocol.

Hard constraints:

- Use ChatGPT image2 / the available `image_gen` tool for master/module image generation.
- Final layout, arrows, alignment, and export must go through local Microsoft Visio controlled by PowerShell COM.
- If Visio is not already open, automatically start `Visio.Application` through COM and set it visible.
- Do not switch the final assembly to draw.io, PowerPoint, browser canvas, matplotlib-only, standalone SVG, or another diagram editor.
- Programmatic PNG/SVG/PDF assembly may be used only as intermediate asset preparation, never as the final deliverable path when the user asked for this skill.
- After generating the overall master figure, stop and ask the user to confirm or give changes. Do not split modules, regenerate modules, or start Visio assembly until the user explicitly approves the overall figure.
- If the user gives changes, regenerate the overall master and ask again. Repeat this loop until the user confirms the master is correct.
- After approval, module generation must 1:1 follow the approved master: same layout, same module proportions, same visual language, same relationship structure.
- Do not use a simple crop of the overall master as the final module unless the user explicitly asks for a pixel-crop reconstruction. The normal workflow is to use the approved master as the reference, then regenerate each module separately with ChatGPT image2 for clarity.
- During final assembly, Visio drawing must happen visibly through PowerShell COM on the user's machine. The user should be able to see Visio open and shapes/assets being placed when practical.

## Core Principles

- **Truth first**: extract real data before drawing.
- **Master first**: generate one full ChatGPT image2 reference figure before splitting modules.
- **Modules next**: split into per-block assets after the master is approved.
- **Real images stay real**: replace hallucinated examples with paper/user-provided real images.
- **Visio assembles**: local Visio handles page size, module placement, arrows, alignment, export, and limited text.
- **Preview is truth**: always inspect the exported preview, not just the editor canvas.
- **Small text is a defect**: replace unreadable notes with icons, keywords, numbers, or summary cards.
- **Preserve reproducibility**: save scripts, assets, preview, and workflow notes.

## End-to-End Workflow

### 1. Build A Figure Specification

Before generating or editing visuals, create a short spec from the paper:

- figure purpose
- module list and order
- exact class/order labels
- counts, equations, totals
- real sample images to use
- annotation type: BBox, mask, class label, keypoint, etc.
- augmentation/operators/method names
- final output target: local Visio `.vsdx` plus Visio-exported preview; SVG/PNG/PDF are intermediate or export companions only

Never let image generation invent exact values. If image2 creates text or numbers, treat them as placeholders until verified.

### 2. Generate A Master Visual

Create one full-width master image first with ChatGPT image2. This is not a Visio draft and not an external-editor composition. It should settle:

- visual style
- layout proportions
- color system
- module hierarchy
- arrow direction
- icon language
- density

For ChatGPT image2 prompts, request:

- conference/paper figure style
- clean flat vector-like infographic
- restrained palette
- no photoreal stock look
- no watermark
- minimal text unless large and essential
- modules without excessive internal arrows

After generating the master, show it to the user and ask for approval. This is a required gate.

Master review loop:

1. Generate the full master.
2. Show the master preview to the user.
3. Ask whether there are layout, content, data, style, readability, or annotation problems.
4. If the user gives changes, revise/regenerate the full master and return to step 2.
5. Only after the user confirms the master is correct, proceed to module generation/splitting.

Do not proceed past this gate on assumption, silence, or partial approval.

### 3. Replace Critical Images With Real Assets

If the figure includes samples, defects, microscopy, sensor output, charts, or experimental screenshots:

- use the real file, not generated substitutes
- crop to the required figure ratio
- draw user/paper-approved annotation boxes
- keep card border/radius/style consistent with the master
- save both raw and annotated versions

If the user wants to draw a box manually, export same-size unannotated cards for them, then re-import their annotated version.

### 4. Regenerate The Approved Master As Clear Modules

Once the master is accepted, decompose it into multiple modules and regenerate each module with ChatGPT image2 using the approved master as the style reference.

The final figure must 1:1 replicate the approved master. Module generation exists to improve clarity at high resolution, not to redesign the figure.

Default module workflow:

1. Identify module boundaries from the approved master.
2. Write a dedicated ChatGPT image2 prompt for each module.
3. Include the approved master style, module role, exact content, and no-cross-module-arrow constraint in each prompt.
4. Generate each module separately with ChatGPT image2.
5. Inspect each module for clarity and fidelity.
6. If a module drifts from the approved master, regenerate that module before Visio assembly.

Use cropping only for:

- extracting a temporary reference,
- preserving a user-approved real sample/card,
- emergency 1:1 reconstruction when image2 cannot follow the module,
- explicit user request.

Do not silently substitute crop-only modules for the normal image2 module regeneration process.

Possible module set:

- header/title if needed
- one asset per major module
- footer if needed
- separate arrows only if they are not to be editor-drawn

Module requirements:

- no cross-module arrows inside generated module assets
- no long unreadable text
- use icons, large keywords, and exact numeric cards
- real paper images must replace generated placeholders
- module content must preserve the approved master layout and visual semantics
- if a regenerated module drifts from the master, revise the module until it matches

Save:

- `module_name.png` for faithful rendering
- `module_name.svg` wrapper or true vector SVG when useful
- a manifest with crop boxes and sizes

Use bundled script `scripts/split_modules.py` only to make reference crops, dimensions, SVG wrappers, or fallback assets. Normal final modules are regenerated with ChatGPT image2 after the master is approved.

### 5. Assemble Only In Local Visio Via COM

Use PowerShell COM to control Visio. Required pattern:

- try `[Runtime.InteropServices.Marshal]::GetActiveObject("Visio.Application")`
- if that fails, create `New-Object -ComObject Visio.Application`
- set `$visio.Visible = $true`
- create/open the target `.vsdx`
- clear only the target page/shapes you own
- import module assets
- place modules with exact coordinates
- draw cross-module arrows in Visio
- export preview with `$page.Export(...)`

Realtime requirement:

- Run the Visio COM script directly on the user's Windows machine.
- Keep `$visio.Visible = $true`.
- Prefer a visible Visio session over hidden/background-only automation.
- If using `Start-Process powershell.exe`, do not hide Visio itself; only the helper PowerShell window may be hidden.
- Do not simulate Visio output by composing SVG/PNG/PDF outside Visio and calling that the final drawing.
- The final deliverables must include a `.vsdx` created/updated through Visio COM and an exported preview from Visio.

Use Visio for:

- page size/aspect ratio
- importing modules
- placing modules at exact coordinates
- drawing cross-module arrows
- final alignment
- preview export

Assembly target:

- The Visio result must visually reproduce the approved master 1:1.
- Use the approved master dimensions as the coordinate reference.
- Place regenerated module assets using exact master-derived coordinates whenever possible.
- Cross-module arrows in Visio must match the approved master positions and style.
- Do not "improve" layout during Visio assembly unless the user explicitly asks for that change.

Do not rebuild complex module internals in Visio if a high-quality module asset already exists.

Important Visio compatibility note: some Visio versions import SVGs containing `<image>` tags but export them blank. If that happens:

- keep SVG wrappers as deliverables,
- use PNG imports for the final visible Visio figure,
- document the compatibility result,
- optionally create a true-vector SVG test version only if image quality remains acceptable.

Do not replace Visio with another editor because SVG import is inconvenient. Use PNG module imports in Visio if that is the only way to preserve 1:1 fidelity.

### 6. Optimize Readability

After exporting the preview, inspect at the target paper scale. Fix:

- tiny explanatory text
- crowded footers
- blurry generated labels
- icons that do not match meaning
- arrows that touch or overlap content
- numbers too small to read

Convert long notes into:

- icon + keyword
- numeric summary cards
- short formula cards
- large labels
- legend-like chips

### 7. Verify And Iterate

Before finalizing:

- export a preview image
- visually inspect module positions, arrows, and text readability
- compare preview against the approved master when applicable
- verify 1:1 reproduction of the approved master layout
- inspect file/package structure if using Visio
- verify paper data one more time
- confirm requested unchanged areas were not modified

See `references/checklists.md` for detailed audit checklists.

## Recommended Output Layout

Use a reproducible structure:

```text
figure_name_master.png
figure_name_preview.png
figure_name.vsdx
figure_name_workflow.md
assets/
  modules/
    header.png
    module_1.png
    module_1.svg
    module_2.png
    module_2.svg
  real_samples/
tools/
  build_master_or_patch.py
  split_modules.py
  rebuild_visio_com.ps1
```

## Useful Bundled Scripts

- `scripts/split_modules.py`: crop an approved master into reference PNGs and SVG wrappers from a JSON manifest; use for coordinates, style references, or fallback assets, not as the normal replacement for image2 module regeneration.
- `scripts/compare_previews.py`: resize two previews to a common size and create a diagnostic diff.

Read or run these scripts as needed; do not paste large code into the conversation unless the user asks.

## Final Response Pattern

When done, report:

- final Visio file path
- preview path
- asset directory
- what changed
- what was intentionally preserved
- any compatibility caveat, such as SVG import behavior

Keep the final answer short and show the preview image when the app supports it.
