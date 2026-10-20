# BibWord for Mac

[English](README_EN.md) | **简体中文**

> **实验版 v0.2**：在 Microsoft Word for Mac 中使用 BibTeX 引用，不需要 Zotero。支持粘贴 BibTeX、导入 `.bib`、通过 citation key 插入引用，以及生成简化的数字编号参考文献。
>
> ⚠️ 此项目尚未在不同版本的 Mac Word 上完成端到端验证。**请在文档副本中测试**，不要直接用于唯一的正式文件。

## 功能

- **粘贴导入**：直接从 Google Scholar / DBLP 等复制完整 BibTeX 条目，粘贴到侧边栏，支持一次导入多条。
- **文件导入**：选择本地 `.bib` 文件；同名 citation key 再次导入会覆盖旧记录。
- **搜索与引用**：按 citation key、标题、作者或年份检索，在光标位置插入 `[1]` 形式的引用。
- **LaTeX 式输入**：在 Word 正文输入 `\cite{smith2024}`，然后点击插件中的「替换文中 `\cite{key}`」。
- **参考文献**：在文末生成参考文献列表，并可在编辑后更新编号。
- **本地运行**：插件页面由本机 `https://localhost:8765` 提供，不依赖在线文献管理账户。

## 系统要求

- macOS、桌面版 Microsoft Word（需支持 Office 加载项侧载和所用 Word JavaScript API）。
- Python 3（运行本地 HTTPS 服务）。
- [Homebrew](https://brew.sh/)（首次安装时用来安装 `mkcert`）。
- Word 支持第三方加载项；某些单位或学校管理的 Word 可能禁用加载项。

> 目前仅提供 **macOS** 安装脚本。Windows 用户不能直接使用该脚本。

## 安装（Mac）

### 1. 获取项目

在 GitHub 页面点击 **Code → Download ZIP** 并解压，或者使用 `git clone <你的仓库地址>`。打开 Terminal（终端），进入包含 `install-mac.sh` 的项目目录：

```bash
cd /path/to/mac-bibword-addon
```

也可以先输入 `cd `（末尾保留空格），把文件夹拖到终端，再按回车。

### 2. 安装 Homebrew（如果还没有）

先检查：

```bash
brew --version
```

如提示 `command not found`，参考 [Homebrew 官网](https://brew.sh/) 安装。Apple Silicon Mac 安装后，可能还需将 `/opt/homebrew/bin` 加入 PATH，具体以 Homebrew 安装器输出为准。

### 3. 安装并启动 BibWord

在项目目录执行：

```bash
bash install-mac.sh
```

脚本会：

1. 检查或通过 Homebrew 安装 `mkcert`；
2. 为 `localhost` 生成并信任本地 HTTPS 证书（可能要求输入 Mac 密码）；
3. 将 `manifest.xml` 复制到 Word 的本地加载项目录；
4. 启动 Python HTTPS 服务器。

成功启动时会显示：

```text
BibWord running at https://localhost:8765
```

**使用插件期间请保持此终端窗口打开。** 按 `Control + C` 可以停止服务。

### 4. 在 Word 中打开加载项

1. 保存并完全退出 Word（`Command + Q`），再重新打开。
2. 打开一个用于测试的 Word 文档。
3. 在 Word 功能区查看 **开始 → 加载项**，或 **插入 → 我的加载项**。
4. 如果能够看到 **BibWord 引用**，点击它，打开右侧侧边栏。

⚠️ 加载项菜单的位置和是否显示取决于 Word 版本及组织策略。**脚本复制清单并不保证 Word 一定会自动显示插件**；如找不到，请先看下面的故障排查。

## 使用教程

### 方式 A：直接粘贴 BibTeX（推荐）

在 BibWord 侧边栏的「① 导入文献库」中粘贴完整条目：

```bibtex
@article{smith2024,
  title   = {Streaming Vision-Language Models},
  author  = {Smith, John and Lee, Alice},
  journal = {Example Journal},
  year    = {2024}
}
```

点击 **导入粘贴内容**。也可以一次性粘贴多个 `@article{...}`、`@inproceedings{...}` 等条目。

### 方式 B：导入已有 `.bib`

在「① 导入文献库」下方点击文件选择框，选择 `references.bib` 即可。重复的 citation key 会覆盖同名条目。

### 在正文中插入引用

**点选插入**：把 Word 光标放在需要的位置，在「② 插入引用」搜索 `smith2024`，点击 **插入引用**。

**使用 LaTeX 写法**：在正文输入：

```text
Existing methods have limitations \cite{smith2024}.
```

然后点击 **替换文中 `\cite{key}`**。插件会尝试将它转换为 `[1]`。

> 目前仅支持单个 citation key 的 `\cite{key}`；**不支持** `\cite{a,b}`，也不支持完整 LaTeX 引用语法。

### 自动生成参考文献

点击「③ 参考文献」中的 **在文末生成参考文献**。之后新增或删除引用，可点击 **更新编号与文献**。

参考文献按正文中首次出现的引用顺序编号，目前使用**简化数字制**，**不等同于严格 IEEE、APA 或 GB/T 7714 样式**。

## 升级 v0.1 → v0.2

1. 保存 Word 文件并关闭 Word。
2. 回到运行旧插件的终端，按 `Control + C` 关闭旧服务。
3. 下载新版本，备份并替换旧项目目录。
4. 在新目录运行 `bash install-mac.sh`。
5. 重启 Word，重新打开 BibWord 侧边栏；如果仍显示旧界面，可尝试重新打开加载项或重启 Word。

当前文献数据通过 Office 文档设置保存，**不同文档并不自动共享一个全局文献库**。升级前请保存文档并备份原始 `.bib` 文件。

## 常见问题

**Q: 终端提示“请先安装 Homebrew”？** 先从 [brew.sh](https://brew.sh/) 安装 Homebrew，确认 `brew --version` 能运行，再重新执行安装脚本。

**Q: 提示 `Address already in use` / 端口 8765 被占用？** 可能旧版服务还在运行。找到运行旧脚本的终端，按 `Control + C`，再启动新版。

**Q: Word 里没有「BibWord 引用」？** 先在 Mac 浏览器访问 `https://localhost:8765`，确认服务是否可访问。再检查清单目录：

```bash
ls -la ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```

确保 Word 已完全退出重启，并查看 **开始/插入 → 加载项 → 我的加载项**。部分 Word 版本、企业策略或侧载机制可能阻止显示。此版本不保证覆盖所有 Word 环境。

**Q: Safari 提示证书问题？** 确认安装脚本完成 `mkcert -install`，不要在未确认来源时绕过 HTTPS 安全警告。

**Q: 为什么导入后其他文档看不到文献？** 当前文献库随 Word 文档设置保存，不是跨文档同步的数据库。请在另一个文档中重新导入 BibTeX。

**Q: 为什么编号或替换没有生效？** 请确认 Word 已打开可编辑文档，citation key 与已导入条目完全一致。当前是实验版本，部分 Word 版本可能存在兼容问题。

## 当前限制

- 仅支持单键 `\cite{key}`，尚不支持多键或复杂引用命令。
- 文献输出为简化格式，不能保证符合正式期刊/学位论文样式。
- 对 `@string` 宏、复杂 BibTeX 表达式和 LaTeX 转义的支持有限。
- 大型 `.bib` 文献库可能受到 Office 文档设置的容量限制。
- 更新、重排及删除后的引用一致性尚未经真实 Mac Word 全面验证；先备份文档。
- 需要本地 HTTPS 服务持续运行；尚无一键安装的签名 `.pkg`/App Store 发布包。

## 隐私说明

当前源码使用 `127.0.0.1:8765` 提供本地页面，不需要登录 Zotero/Paperpile，文献数据保存在当前 Word 文档的 Office 设置中。**但如果文档上传到 OneDrive 等云端，文档中的文献数据也可能随文档同步**。请根据实际使用环境自行评估隐私需求。

## 反馈与贡献

欢迎通过 GitHub **Issues** 报告问题，请尽量提供 macOS 版本、Word 版本、错误截图以及复现步骤（不要上传包含敏感信息的私人文档）。也欢迎提交 Pull Requests。

## 许可

**尚未指定开源许可证。** 如果希望允许他人合法地复制、修改和再分发代码，请在发布仓库前添加合适的 `LICENSE` 文件（例如 MIT），并确认项目内所有代码与素材均可按所选许可发布。
