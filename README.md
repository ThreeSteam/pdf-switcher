# PDF Switcher

Windows WPF Word-to-PDF converter. A portable executable with embedded PowerShell workers; desktop Microsoft Word is required. Files are processed locally and exported beside their source with the same base name.

Build: run `powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1` on Windows with .NET Framework 4.x. Output: `portable\PDF Switcher.exe`.

The UI supports multi-file/folder drops, duplicate filtering, a serial queue, progress, cancel remaining, retries, and locating results. Folder drops include immediate children only. Existing PDFs are skipped by default. Explicit overwrite exports to a temporary file then replaces the existing file atomically. Source documents are opened read-only and closed without saving.

Revision modes: follow the view Word establishes when opening the saved document, clean export, or all revisions/comments. This does not capture unsaved state from an already open Word window.

Each conversion owns a separate hidden Word process, identified by process ID and start time. The UI remains responsive and enforces a three-minute per-phase timeout. A separate process per document favors fault isolation over maximum batch speed.

Developer smoke hooks: `--preview <png-path>` renders the WPF window; `--test <folder> [auto|clean|markup|overwrite]` uses the production queue and writes `result.json` beside test files. Use only a dedicated test folder.

The reference styling and native Word attachment helper originate from the user's local Word Redline project. No client documents are included in the portable distribution.
