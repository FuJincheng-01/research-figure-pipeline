# Research Figure Pipeline

**Language / 语言**: [中文](#中文说明) | [English](#english-readme)

---

<a id="中文说明"></a>

# Research Figure Pipeline（科研绘图插件）

一个用于绘制论文级科研图的 Codex 插件和 skill。它采用严格的两阶段流程：

1. 先使用 ChatGPT image2 生成整张参考图，并让用户审核确认；
2. 用户确认整体方案后，再把整图拆成多个清晰模块分别重新生图；
3. 最后通过 Windows 本机 Microsoft Visio + PowerShell COM 完成可编辑 `.vsdx` 组装、箭头连接、对齐和导出。

这个项目来自 OmniDet 数据集构建图的真实绘制流程，适合 CVPR / ICCV / TIP / IEEE 风格的数据集构建图、数据增强图、方法流程图、模型结构图和实验 pipeline 图。它的核心思路是：AI 生图负责把每个模块画得好看，Visio 负责最终排版、箭头关系和论文可编辑文件。

## 重要限制

这不是一个独立的通用生图程序。

- 主要支持在 Codex 中使用，因为它依赖 Codex 的 ChatGPT image2 生图能力。
- 最终组装需要 Windows + Microsoft Visio。
- 最终布局、对齐、箭头和导出必须通过本机 Visio + PowerShell COM 完成。
- draw.io、PowerPoint、浏览器 canvas、matplotlib、单独 SVG/PNG/PDF 拼图只能作为中间准备步骤，不能作为最终组装路径。
- 如果 Visio 没有打开，流程会要求通过 COM 自动启动 `Visio.Application`，并设置为可见。

## 仓库内容

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

`.codex-plugin` 是 Codex 插件封装，`skills/research-figure-pipeline` 是可单独安装的 skill。

## 核心流程

这个 skill 会强制执行下面的绘图顺序：

1. 阅读论文、PDF、LaTeX 源码或手稿材料。
2. 先提取真实内容：标签、数量、类别顺序、公式、标注类型、方法名、总数和真实样例图。
3. 使用 ChatGPT image2 生成一张完整 master 图。
4. 把 master 图展示给用户，并等待用户明确确认。
5. 如果用户提出修改意见，就重新生成整张 master 图并再次确认。
6. 只有在用户确认 master 图无误后，才拆解成多个模块。
7. 每个模块再单独调用 ChatGPT image2 重新生成，提升清晰度。
8. 对于缺陷图、实验截图、样本图等关键内容，使用论文或用户提供的真实图片替换生成图。
9. 使用本机 Microsoft Visio + PowerShell COM 组装所有模块。
10. 跨模块箭头、连接关系、对齐和最终导出都在 Visio 中完成。
11. 从 Visio 导出预览图，并检查是否和用户确认过的 master 图 1:1 对齐、文字是否清晰。

核心规则：

```text
ChatGPT image2 负责生成视觉模块；Visio 负责组装模块关系。
```

## 快速安装为 Codex 插件

克隆仓库并运行安装脚本：

```powershell
git clone https://github.com/FuJincheng-01/research-figure-pipeline.git
cd research-figure-pipeline
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

安装脚本会：

- 在需要时把插件复制到 `%USERPROFILE%\plugins\research-figure-pipeline`；
- 创建或更新 `%USERPROFILE%\.agents\plugins\marketplace.json`；
- 把插件加入 personal marketplace；
- 如果能找到 Codex CLI，则自动执行 `codex plugin add research-figure-pipeline@personal`。

安装后请新开一个 Codex 对话，让 Codex 加载新的插件和 skill 元数据。

## 手动安装为 Codex 插件

把仓库克隆到本地插件目录：

```powershell
cd $env:USERPROFILE\plugins
git clone https://github.com/FuJincheng-01/research-figure-pipeline.git
```

确认 personal marketplace 文件中存在下面这个条目：

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

然后执行：

```powershell
codex plugin add research-figure-pipeline@personal
```

安装后请新开一个 Codex 对话。

## 单独安装为 Skill

如果不需要插件封装，只想安装 skill，可以复制 skill 文件夹：

```powershell
Copy-Item `
  "$env:USERPROFILE\plugins\research-figure-pipeline\skills\research-figure-pipeline" `
  "$env:USERPROFILE\.codex\skills\research-figure-pipeline" `
  -Recurse -Force
```

然后在新的 Codex 对话中调用：

```text
Use $research-figure-pipeline to create a publication-ready research figure from my paper materials.
```

## 示例 Prompt

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

## 典型输出

一次完整绘图通常会产出：

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

## Visio COM 要求

最终组装应该使用 PowerShell COM，基本模式如下：

```powershell
try {
  $visio = [Runtime.InteropServices.Marshal]::GetActiveObject("Visio.Application")
} catch {
  $visio = New-Object -ComObject Visio.Application
}
$visio.Visible = $true
```

用户应能看到 Visio 打开并执行组装。最终 `.vsdx` 和预览图应该来自 Visio，而不是来自外部 PNG/SVG 拼图程序。

## 关于 SVG 和 PNG

部分 Visio 版本可以导入包含嵌入 PNG 的 SVG，但导出时会变空白。遇到这种情况时，可以保留 SVG wrapper 作为附属资产，但最终 Visio 图中建议导入 PNG 模块以保证视觉一致性。

## 许可证

MIT License。见 [LICENSE](LICENSE)。

---

# English README

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
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
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
