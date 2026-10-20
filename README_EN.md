# BibWord for Mac

**English** | [简体中文](README.md)

> **Experimental v0.2**: Add BibTeX-based citations to Microsoft Word for Mac without Zotero. Paste BibTeX, import `.bib` files, insert citations by citation key, and generate a simple numbered bibliography.
>
> ⚠️ This project has **not** been fully tested end to end across real Mac Word environments. **Use a copy of your document for testing**, not your only production copy.

## Features

- **Paste BibTeX:** Copy full BibTeX entries from Google Scholar, DBLP, or another source and paste one or more entries into the task pane.
- **Import `.bib` files:** Load an existing bibliography file; imported entries with the same citation key overwrite previous entries.
- **Search and cite:** Find entries by key, title, author, or year and insert numbered citations such as `[1]` at the Word cursor.
- **LaTeX-like citation keys:** Type `\cite{smith2024}` in Word and click **替换文中 `\cite{key}`** (Replace in-document citation commands).
- **Generate references:** Create a numbered bibliography at the end of the document and refresh numbering as you edit.
- **Local operation:** The add-in is served by `https://localhost:8765`; no Zotero/Paperpile account is required.

## Requirements

- macOS and desktop Microsoft Word with support for Office add-in sideloading and the required Word JavaScript APIs.
- Python 3 for the local HTTPS server.
- [Homebrew](https://brew.sh/) to install `mkcert` if it is not already installed.
- Permission to use Office add-ins (some managed school/company installations disable them).

> Only a **macOS** installation script is currently provided. Do not run it on Windows.

## Installation (macOS)

### 1. Get the source

On GitHub, select **Code → Download ZIP** and extract it, or run `git clone <your-repository-URL>`. Open Terminal and move into the directory containing `install-mac.sh`:

```bash
cd /path/to/mac-bibword-addon
```

Tip: Type `cd ` (with a trailing space), drag the folder into Terminal, and press Enter.

### 2. Install Homebrew if necessary

Check whether it is already available:

```bash
brew --version
```

If not installed, follow the instructions at [brew.sh](https://brew.sh/). On Apple Silicon, you may need to add `/opt/homebrew/bin` to your PATH as instructed by the installer.

### 3. Install and run BibWord

```bash
bash install-mac.sh
```

The script:

1. Installs `mkcert` through Homebrew if needed.
2. Creates and trusts a local HTTPS certificate for `localhost` (macOS may ask for your password).
3. Copies `manifest.xml` into Word's local add-in manifest directory.
4. Starts the Python HTTPS server.

A running server prints:

```text
BibWord running at https://localhost:8765
```

**Keep this terminal open while using the add-in.** Press `Control + C` there to stop the service.

### 4. Open it in Microsoft Word

1. Save your work, quit Word completely (`Command + Q`), and reopen it.
2. Open a test Word document.
3. Look for **Home → Add-ins** or **Insert → My Add-ins** in the Word ribbon.
4. If **BibWord 引用** (BibWord Citations) appears, select it to open the task pane.

⚠️ The exact menu and whether the add-in appears depend on your Word version and organization policies. **Copying the manifest does not guarantee successful sideloading** in every Word setup. See Troubleshooting if it does not appear.

## How to use

### Option A: Paste BibTeX (recommended)

In the add-in pane, find **① 导入文献库** (Import bibliography). Paste one or more complete entries:

```bibtex
@article{smith2024,
  title   = {Streaming Vision-Language Models},
  author  = {Smith, John and Lee, Alice},
  journal = {Example Journal},
  year    = {2024}
}
```

Click **导入粘贴内容** (Import pasted content). You can paste multiple `@article{...}`, `@inproceedings{...}`, etc. entries at once.

### Option B: Import a `.bib` file

Use the file picker under **① 导入文献库** to select `references.bib`. Entries with duplicate citation keys overwrite the previous entries.

### Insert a citation in Word

**Search and insert:** Place the cursor in Word, search for `smith2024` under **② 插入引用** (Insert citation), and click **插入引用** (Insert citation).

**Type a LaTeX-like command:** Enter the following in Word:

```text
Existing methods have limitations \cite{smith2024}.
```

Then click **替换文中 `\cite{key}`**. The add-in attempts to turn it into `[1]`.

> Only `\cite{key}` with **one key** is supported at present. `\cite{a,b}` and full LaTeX citation semantics are **not supported**.

### Generate a bibliography

Under **③ 参考文献** (References), click **在文末生成参考文献** (Generate bibliography at document end). After modifying citations, click **更新编号与文献** (Refresh numbers and references).

Entries are numbered in order of first appearance in the text. Formatting is a **simplified numeric style**, **not** compliant IEEE, APA, or GB/T 7714 output.

## Upgrading from v0.1

1. Save your Word document and quit Word.
2. In the terminal running the old service, press `Control + C`.
3. Download v0.2, back up your old source folder, and replace it with the new one.
4. Run `bash install-mac.sh` from the new directory.
5. Reopen Word and the BibWord task pane. If it still shows an old interface, close/reopen the task pane or restart Word.

The bibliography library is currently stored in the **settings of the current Word document**. Different Word documents do **not** automatically share a global bibliography. Back up the document and your original `.bib` file before upgrading.

## Troubleshooting

**“Please install Homebrew”**: Install Homebrew from [brew.sh](https://brew.sh/), verify `brew --version`, and rerun the script.

**`Address already in use` / port 8765 is busy**: The previous BibWord service may still be running. Stop it using `Control + C` in the old terminal before launching the new one.

**BibWord does not appear in Word**: First visit `https://localhost:8765` in a Mac browser to check that the server is responding. Then check the manifest directory:

```bash
ls -la ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```

Quit and restart Word, and check **Home / Insert → Add-ins → My Add-ins**. Word versions, administrator policies, and sideloading limitations vary; appearance is not guaranteed on every installation.

**Certificate warning in the browser**: Confirm `mkcert -install` completed successfully. Do not bypass certificate warnings for unverified services.

**References are missing from another document**: BibWord stores its library in the current Word document's settings, rather than in a shared database. Import the BibTeX again in the other document.

**Insertion or refreshing fails**: Make sure you have an editable Word document open and the citation key matches a loaded entry exactly. This release is experimental and may have Word-version-specific issues.

## Known limitations

- Single-key `\cite{key}` only; multi-key commands and advanced LaTeX citation commands are unsupported.
- Simplified bibliography style only; not journal-ready IEEE/APA/GB/T 7714.
- Limited support for BibTeX `@string` macros, complex expressions, and LaTeX character escaping.
- Large `.bib` libraries may exceed Office document settings size limits.
- Citation updating, ordering, and deletion behaviors have not been comprehensively verified in real Mac Word installations. Back up important documents.
- The localhost HTTPS server must remain running; no signed `.pkg` or App Store distribution is provided.

## Privacy

The current source serves the add-in locally on `127.0.0.1:8765` and does not require a remote bibliography account. Imported bibliography entries are saved into the active Word document's Office settings. **If your document is synced to OneDrive or other cloud storage, this embedded bibliography data may be synced with it.** Evaluate the privacy implications for your workflow.

## Feedback and contributions

Please use GitHub **Issues** for bug reports and include your macOS version, Word version, error screenshots, and reproduction steps. Do not share private or sensitive documents. Pull requests are welcome.

## License

**No open-source license has been selected yet.** If you want others to be legally permitted to copy, modify, and redistribute the code, add an appropriate `LICENSE` file (for example, MIT) before publishing, and confirm that all included code and assets can be distributed under your chosen terms.
