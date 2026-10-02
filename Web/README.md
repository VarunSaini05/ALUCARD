# ALUCARD Website

This is a dependency-free static download site. It reads version, publication date, and file size from the latest public GitHub Release and only enables downloads after finding an asset named `ALUCARD-Setup.exe`.

## Run locally

From the repository root, start Python's static web server:

```powershell
py -m http.server 8000 --directory Web
```

On Linux or macOS, use:

```sh
python3 -m http.server 8000 --directory Web
```

Open <http://localhost:8000>. No package installation or build step is required. The release request uses the public GitHub API, so the download state reflects the repository's current public release availability.

## Website deployment

`.github/workflows/deploy-website.yml` deploys the contents of `Web/` to GitHub Pages when `Web/` changes on `main`, or when manually dispatched. In the repository, open **Settings > Pages** and set the build and deployment source to **GitHub Actions**. The project site URL is:

<https://varunsaini05.github.io/ALUCARD/>

This repository is public so visitors can read release metadata and download public assets without credentials. The site does not contain a token and cannot download private release assets for anonymous visitors.

## Publish an installer release

The Windows build workflow uses the existing self-hosted Windows x64 runner and its `SOLIDWORKS_INTEROP_DIR` configuration. After the installer builds successfully, it uploads the Actions artifact. On a `v*` tag push, it also creates or updates the matching GitHub Release and uploads `Release/ALUCARD-Setup.exe` as a release asset.

For example, from the repository root:

```sh
git tag v1.0.0
git push origin v1.0.0
```

Use the next appropriate version tag for each release. The site's GitHub API request checks the latest public release for the exact installer asset before enabling the download button. Its stable download route is:

<https://github.com/VarunSaini05/ALUCARD/releases/latest/download/ALUCARD-Setup.exe>

The release automation only runs after the Windows installer job succeeds. The SolidWorks interop assemblies remain local runner inputs and are not uploaded in the installer or committed to the repository.