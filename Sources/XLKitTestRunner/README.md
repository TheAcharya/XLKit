# XLKitTestRunner

A modular test runner for XLKit that generates Excel files for testing and demonstration purposes.

## Structure

```
Sources/XLKitTestRunner/
├── main.swift                    # Entry point with command-line interface
├── ExcelGenerators.swift         # Excel generation functions
├── ImageEmbedGenerators.swift    # Image embedding test functions
├── Templates/                    # Template files for new tests
│   └── TestGeneratorTemplate.swift
└── README.md                     # This file
```

## Usage

### Command Line Interface

```bash
# Run the default test (no-embeds)
swift run XLKitTestRunner

# Run specific test types
swift run XLKitTestRunner no-embeds
swift run XLKitTestRunner embed
swift run XLKitTestRunner comprehensive
swift run XLKitTestRunner security-demo

# Show help
swift run XLKitTestRunner help
```

### Available Test Types

| Test Type | Description | Status |
|-----------|-------------|--------|
| `no-embeds` / `no-images` | Generate Excel from CSV without images | ✅ Implemented |
| `embed` / `with-embeds` / `with-images` | Generate Excel with image embeds | ✅ Implemented |
| `comprehensive` / `demo` | Comprehensive API demonstration with all features | ✅ Implemented |
| `security-demo` / `security` | Demonstrate file path security restrictions | ✅ Implemented |

## Test Descriptions

### no-embeds / no-images
Generates an Excel file from CSV data without any image embeddings. Tests basic CSV import functionality and Excel file generation.

**Output**: `Test-Workflows/Embed-Test.xlsx`

### embed / with-embeds / with-images
Generates an Excel file from CSV data with embedded images. Tests image embedding functionality with perfect aspect ratio preservation.

**Output**: `Test-Workflows/Embed-Test-Embed.xlsx`

**Features**:
- CSV data import with 22 columns and 4 data rows
- Image embedding for all 4 rows using PNG files
- Automatic column width optimization
- Bold headers with gray background
- Perfect aspect ratio preservation for all images
- CoreXLSX validation for Excel compliance

### comprehensive / demo
Comprehensive API demonstration showcasing all XLKit features including multiple sheets, formulas, formatting, and image embedding.

**Output**: `Test-Workflows/Comprehensive-Demo.xlsx`

**Features**:
- Multiple sheets with different content types
- Cell formatting and styling
- Formulas and calculations
- Image embedding with various formats
- Range operations and cell merging
- Complete API demonstration

### security-demo / security
Demonstrates file path security restrictions and security features of XLKit.

**Features**:
- File path validation
- Security logging demonstration
- Rate limiting examples
- Input validation testing
- Security feature showcase

## Adding New Tests

### 1. Use the Template

Copy the template file:
```bash
cp Sources/XLKitTestRunner/Templates/TestGeneratorTemplate.swift Sources/XLKitTestRunner/YourTestName.swift
```

### 2. Modify the Template

1. Rename the function to match your test
2. Update configuration (test name, output file)
3. Implement your test logic
4. Add error handling and security logging

### 3. Register in main.swift

Add your test to the switch statement in `main.swift`:

```swift
case "your-test-name":
    print("Executing: Your Test Description")
    try yourTestName()
    validateExcelFile("Test-Workflows/Your-Test.xlsx")
```

### 4. Update Help Text

Add your test to the help function in `main.swift`.

### 5. Create GitHub Actions Workflow (Optional)

If your test needs to run in CI/CD, create a new workflow file.

## Naming Conventions

### Function Names
- Use camelCase: `generateExcelWithImages()`
- Be descriptive: `testConditionalFormatting()`
- Use verbs: `generate`, `test`, `validate`

### Test Type Names
- Use kebab-case: `with-images`, `csv-import`
- Be concise but clear
- Support aliases: `no-embeds` and `no-images`

### File Names
- Use PascalCase: `ExcelGenerators.swift`
- Match function names: `CellFormattingTests.swift`

## Output Structure

All generated Excel files are saved to:
```
Test-Workflows/
├── Embed-Test.xlsx          # From no-embeds test
├── Embed-Test-Embed.xlsx    # From embed test (with images)
├── Comprehensive-Demo.xlsx  # From comprehensive test
└── [Your-Test].xlsx         # From your custom tests
```

## Security Features

All test functions include comprehensive security features:

### Security Logging
- All operations are logged with structured data
- Timestamps and operation details captured
- Audit trail maintained for all file operations

### Rate Limiting
- Prevents test abuse and resource exhaustion
- Configurable limits for different operations
- Automatic enforcement across all tests

### Input Validation
- All test inputs validated for security
- File path sanitization and validation
- Image format and size validation

### File Quarantine
- Suspicious files automatically quarantined
- Pattern detection for malicious content
- Size and format validation

### Checksum Verification
- Optional file integrity verification
- SHA-256 hashes for file authenticity
- Tamper detection capabilities

## Error Handling

All test functions should:
- Use descriptive error messages
- Exit with appropriate codes (0 for success, 1 for failure)
- Log progress with `[INFO]`, `[SUCCESS]`, `[ERROR]` prefixes
- Include security logging for all operations

## Dependencies

- XLKit: Main library for Excel generation
- CoreXLSX: Excel file validation and compliance checking
- Foundation: File system operations and utilities

## Validation

Every generated Excel file is automatically validated using CoreXLSX to ensure:
- Full OpenXML compliance and Excel compatibility
- Proper file structure and relationships
- Image embedding integrity with perfect aspect ratio preservation
- Cell and row data accuracy
- Professional-quality exports for all video and cinema formats