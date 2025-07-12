# Changelog

### 1.0.3

**Released:**
- 12th July 2025

**🔧 Improvements:**
- Identified and resolved scaling inconsistencies between XLKit test and MarkersExtractor integration
- Fixed XLKit test to use default parameters instead of manual overrides for consistent behavior
- Established consistent API usage pattern: let XLKit handle sizing automatically
- Documented scale parameter options and best practices for image embeddings
- Updated documentation (AGENT.MD and .cursorrules) with scaling API details and integration guidelines

---

### 1.0.2

**🎉 Released:**
- 9th July 2025

**🔧 Improvements:**
- Updated Security defaults

---

### 1.0.1

**🎉 Released:**
- 9th July 2025

**🔧 Improvements:**
- Added SecurityManager with rate limiting, security logging, file quarantine, and checksums
- Integrated security features throughout XLSXEngine, ImageUtils, and XLKit API
- Added comprehensive input validation for all user inputs
- Replaced system zip command with pure Swift ZIP library to eliminate command injection risk
- Fixed Swift 6.0 concurrency issues with @MainActor and throws propagation
- Enhanced XLKitTestRunner with security integration and better error handling
- Updated documentation (AGENT.MD and .cursorrules) with security features
- Fixed test runner hanging issues and build errors
- Expanded from 9 to all 17 professional video and cinema aspect ratios with pixel-perfect preservation
- Added comprehensive security requirements and validation processes to development guidelines

---

### 1.0.0

**🎉 Released:**
- 7th July 2025

This is the first public release of **XLKit**!
