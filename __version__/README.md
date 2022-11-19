# Updating procedure perform step 1 and 2

## Step 1: On main.py
1. Edit variable `__version__=` to the new version tag.
2. Create setup.exe with corsponding version.

## Step 2: On github
 1. Create a release with setup.exe as assets and copy the link of the asset
 2. Open `Overwatch-Server-Selector/__version__/__latestversion__.txt`
 3. First line corresponds to the latest version matches latest main.py `__version__=`.
 4. Second line is the url copied from asset to download the version.
