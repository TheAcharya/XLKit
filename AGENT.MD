# AGENT.MD

## Project: XLKit

**XLKit** is a pure Swift library for creating and writing Microsoft Excel (.xlsx) files on macOS.  
There is no UI—this is a backend/data library only.

---

## Agent Guidelines

### 1. Scope of Work
- Focus on robust, type-safe, and **extremely easy-to-use** Swift code.
- All code must be compatible with Swift 6.0 and macOS 13+.
- No UI code; only library logic, tests, and documentation.
- Support for advanced Excel features (images, GIFs, formulas, merging, column sizing, etc.) is encouraged, but must be implemented in a modular, maintainable, and user-friendly way.

### 2. Coding Standards
- Use modern Swift best practices (async/await, error handling, value types where possible).
- All public APIs must be documented with doc comments and designed for discoverability and ease of use.
- All new features must include corresponding unit tests and usage examples.
- Avoid global state and side effects.
- Prefer fluent, chainable, and bulk APIs where possible.

### 3. Testing
- All code must be covered by unit tests in `Tests/XLKitTests`.
- Tests should be clear, deterministic, and not rely on external resources.
- All new features (including image and column sizing) must be tested.

### 4. Documentation
- Update `README.md` with new features or API changes.
- All public types and methods must have doc comments and usage examples.

### 5. CI/CD
- Ensure the GitHub Actions workflow (`.github/workflows/build.yml`) passes before merging.
- No platform support except macOS.

### 6. File/Folder Structure
- Source code: `Sources/XLKit/`
- Tests: `Tests/XLKitTests/`
- CI: `.github/workflows/`
- No UI, assets, or platform-specific code outside macOS.

---

## Out of Scope

- No UI or app code.
- No Windows or Linux support.
- No dependencies except Swift standard library and Foundation.

---

## Agent Behavior

- Always prefer clarity, safety, and **ease of use** over brevity.
- If unsure, ask for clarification in a comment or PR.
- Do not auto-merge breaking changes. 