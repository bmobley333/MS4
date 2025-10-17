# Flex Project Architecture Summary

This document outlines the architecture for the Flex TTRPG project, which is built on Google Sheets and powered by Google Apps Script.

---

## Core Philosophy

- **Central Library (FlexLib):** The vast majority of game logic resides in a single `FlexLib` script project. Player-facing sheets contain only minimal "trigger" code that calls this central library.
- **Tag-Driven Architecture:** This is the most critical paradigm. The system never relies on hardcoded column letters, row numbers, or sheet order. All data is accessed through row and column tags (e.g., `colTags.abilityname`). This makes the system robust, flexible, and resilient to layout changes.
- **Robust User Experience (UX):** The system prioritizes clear, non-intrusive feedback. Functions provide immediate toast notifications (`fShowToast`) to confirm they are running, and conclude with clear success or error messages (`fShowMessage`).

---

## Key Architectural Patterns

- **Performance via Caching (`fGetSheetData`):** All spreadsheet data reads are funneled through the `fGetSheetData` gatekeeper function. It uses a session cache by default for performance and supports a `forceRefresh` option for reading the absolute latest user input when necessary.
- **System Resilience (Self-Healing):** The system is designed to recover from common user errors.
  - Folder access uses `fGetSubFolder` to find critical folders even if the main project folder is moved or renamed.
  - File access uses `fGetVerifiedLocalFile` to automatically restore master templates if a player accidentally deletes them.
- **Data Segregation:** For complex user-generated content (like custom powers), we use a two-sheet system:
  1.  A player-facing "working" sheet (e.g., `<Powers>`) where they input and edit data.
  2.  A hidden, "published" sheet (e.g., `<VerifiedPowers>`) that contains only the clean, validated data. The rest of the game system reads exclusively from this clean sheet to ensure data integrity.

---

## Development & Implementation

- **Development Environment:** The entire project is managed locally in VS Code and synced with Google's servers using `clasp`. Version control is handled through Git/GitHub.
- **Global `g` Object:** A global constants object (`g`) in `FlexLib` serves as the single source of truth for IDs, version numbers, and other constants.
- **Code Organization:** `FlexLib` is organized into multiple logical files (e.g., `Menus.js`, `Powers.js`, `Custom.js`) for clarity.
- **Notation:** The notation `<SheetName>` is used in discussions as a shortcut for "the sheet named 'SheetName'".
