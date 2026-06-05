# Research Figure Pipeline

A Codex plugin and skill for building publication-style research figures with a strict two-stage workflow:

1. generate and approve a full reference figure with ChatGPT image2,
2. regenerate the approved figure as clear modules,
3. assemble the final editable/exportable figure in Microsoft Visio through PowerShell COM.

This project was distilled from an OmniDet dataset-construction figure workflow for computer-vision paper figures. It is useful for CVPR/ICCV/TIP/IEEE-style dataset, augmentation, method, pipeline, and architecture diagrams where AI-generated visual modules look better than hand-drawn editor shapes, but final alignment and arrows still need a real diagram editor.

## Important Limitations

This is not a standalone image-generation program.

- It is designed for Codex because it relies on Codex's ChatGPT image2 image-generation capability.
- It requires Microsoft Visio on Windows for the final assembly path.
- Final layout, alignment, arrows, and export must be done through local Visio controlled by PowerShell COM.
- Other tools such as draw.io, PowerPoint, browser canvas, matplotlib, or standalone SVG/PNG/PDF composition are allowed only as intermediate preparation, not as the final figure assembly path.
- If Visio is not open, the workflow should start `Visio.Application` automatically through COM and set it visible.

## What Is Included

```text
.codex-plugin/
  plugin.json
skills/
  research-figure-pipeline/
    SKILL.md
    agents/openai.yaml
    references/checklists.md
    scripts/split_modules.py
    scripts/compare_previews.py
```

The plugin wraps the skill for Codex plugin usage. The skill can also be installed standalone.

## Core Workflow

The skill enforces this sequence:

1. Read the paper, PDF, LaTeX source, or manuscript material.
2. Extract factual figure content first: labels, counts, class order, equations, annotation type, operators, totals, and real sample images.
3. Generate one full-width master visual with ChatGPT image2.
4. Show the master to the user and wait for explicit approval.
5. If the user requests changes, regenerate the full master and ask again.
6. Only after approval, decompose the approved master into modules.
7. Regenerate each module separately with ChatGPT image2 for higher clarity.
8. Replace critical generated samples with real paper/user-provided images when needed.
9. Assemble the regenerated modules in local Microsoft Visio through PowerShell COM.
10. Draw cross-module arrows in Visio.
11. Export a preview from Visio and verify readability and 1:1 fidelity against the approved master.

The central rule is:

```text
ChatGPT image2 builds visual modules; Visio assembles relationships.
```

## Quick Install As A Codex Plugin

Clone this repository and run the installer:

```powershell
git clone https://github.com/FuJincheng-01/research-figure-pipeline.git
cd research-figure-pipeline
.\install.ps1
```

The installer will:

- copy the plugin to `%USERPROFILE%\plugins\research-figure-pipeline` when needed,
- create or update `%USERPROFILE%\.agents\plugins\marketplace.json`,
- add the plugin entry to the personal marketplace,
- run `codex plugin add research-figure-pipeline@personal` if the Codex CLI is available.

Open a new Codex thread after installation so the plugin and skill metadata are loaded.

## Manual Install As A Codex Plugin

Clone this repository into your local plugin directory:

```powershell
cd $env:USERPROFILE\plugins
git clone https://github.com/FuJincheng-01/research-figure-pipeline.git
```

Make sure your personal Codex marketplace has an entry like this in:

```text
%USERPROFILE%\.agents\plugins\marketplace.json
```

```json
{
  "name": "research-figure-pipeline",
  "source": {
    "source": "local",
    "path": "./plugins/research-figure-pipeline"
  },
  "policy": {
    "installation": "AVAILABLE",
    "authentication": "ON_INSTALL"
  },
  "category": "Productivity"
}
```

Then install from the personal marketplace:

```powershell
codex plugin add research-figure-pipeline@personal
```

Open a new Codex thread so the updated plugin and skill metadata are loaded.

## Install As A Standalone Skill

If you only want the skill without the plugin wrapper, copy the skill folder:

```powershell
Copy-Item `
  "$env:USERPROFILE\plugins\research-figure-pipeline\skills\research-figure-pipeline" `
  "$env:USERPROFILE\.codex\skills\research-figure-pipeline" `
  -Recurse -Force
```

Then open a new Codex thread and invoke:

```text
Use $research-figure-pipeline to create a publication-ready research figure from my paper materials.
```

## Example Prompts

```text
Use $research-figure-pipeline to draw a dataset construction figure for this paper.
First generate a full master image with ChatGPT image2 and wait for my approval.
After I approve it, regenerate each module separately and assemble the final VSDX through visible Visio COM.
```

```text
Use $research-figure-pipeline to convert this method section into a CVPR-style pipeline figure.
Use real sample images where I provide them.
Final output must be a Visio .vsdx plus a Visio-exported preview.
```

## Expected Outputs

A typical run should produce:

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
  split_modules.py
  rebuild_visio_com.ps1
```

## Visio COM Requirement

The final assembly should use PowerShell COM with the following pattern:

```powershell
try {
  $visio = [Runtime.InteropServices.Marshal]::GetActiveObject("Visio.Application")
} catch {
  $visio = New-Object -ComObject Visio.Application
}
$visio.Visible = $true
```

The user should be able to see Visio open and the figure being assembled when practical. The final `.vsdx` and preview should come from Visio, not from an external raster/SVG compositor.

## Notes On SVG And PNG

Some Visio versions import SVG files containing embedded PNG images but export them blank. If that happens, keep SVG wrappers as supporting assets, but use PNG module imports for the final visible Visio figure. This preserves the 1:1 visual fidelity of the approved master.

## License

MIT License. See [LICENSE](LICENSE).
