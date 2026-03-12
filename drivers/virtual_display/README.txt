Place bundled virtual display driver files here for one-click installation.

Required:
- At least one .inf file in this folder (or any subfolder)
- All files referenced by the .inf (cat, sys, dll, etc.)

AutoPlua will:
1) Prefer a user-selected INF in the UI if provided.
2) Otherwise scan this folder recursively and pick a suitable INF automatically.

Packaging note:
- Include the full drivers/virtual_display directory in your installer or release package.
- Installation still requires Administrator privileges on Windows.
