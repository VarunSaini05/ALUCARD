# ALUCARD

ALUCARD is a SolidWorks automation project for automated Cut List to DXF extraction.

## Project Status

The current implementation is a working SolidWorks VBA automation prototype. The existing macro will be added to this repository separately.

## Current Workflow

The automation workflow:

1. Read the SolidWorks Cut List.
2. Extract individual bodies.
3. Create individual SLDPRT files.
4. Convert parts to Sheet Metal.
5. Flatten parts.
6. Export DXF flat patterns.
7. Recover failed parts using Insert Bends.
8. Use a largest-planar-face fallback where appropriate.
9. Report parts that remain unresolved and require manual work.

## SolidWorks Interface Goal

The intended future SolidWorks add-in interface is:

**ALUCARD → Extract DXF**

## Repository Structure

- `SolidWorks/VBA` contains the proven VBA automation prototype.
- `Addin` is for a future SolidWorks add-in implementation.
- `Web` is reserved for a potential future public web application.
- `Docs` contains project and development documentation.
- `Tests` contains future tests for supported project functionality.

## Development Environment

GitHub Codespaces is currently used for development, documentation, and version control. SolidWorks runs separately on a Windows workstation; it is not installed in Codespaces.

## Roadmap

- **Phase 1:** Working VBA automation.
- **Phase 2:** GitHub version control and documentation.
- **Phase 3:** SolidWorks ALUCARD add-in with an "Extract DXF" command.
- **Phase 4:** Testing and packaging.
- **Phase 5:** Potential public web application.

## Important Development Rule

The existing working VBA macro is the source of truth for the engineering workflow. Do not modify its engineering logic unless explicitly requested. The add-in and any future web application must not replace or change that workflow without explicit direction.
