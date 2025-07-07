# XLKitTestRunner

A modular test runner for XLKit that generates Excel files for testing and demonstration purposes.

## Structure

```
Sources/XLKitTestRunner/
├── main.swift                    # Entry point with command-line interface
├── ExcelGenerators.swift         # Excel generation functions
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
swift run XLKitTestRunner with-images
swift run XLKitTestRunner csv-import
swift run XLKitTestRunner formatting

# Show help
swift run XLKitTestRunner help
```

### Available Test Types

| Test Type | Description | Status |
|-----------|-------------|--------|
| `no-embeds` / `no-images` | Generate Excel from CSV without images | Implemented |
| `with-embeds` / `with-images` | Generate Excel with image embeds | Future |
| `csv-import` | Test CSV import functionality | Future |
| `formatting` | Test cell formatting features | Future |

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
4. Add error handling

### 3. Register in main.swift

Add your test to the switch statement in `main.swift`:

```swift
case "your-test-name":
    print("Executing: Your Test Description")
    yourTestName()
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
├── Template-Test.xlsx       # From template
└── [Your-Test].xlsx         # From your custom tests
```

## Error Handling

All test functions should:
- Use descriptive error messages
- Exit with appropriate codes (0 for success, 1 for failure)
- Log progress with `[INFO]`, `[SUCCESS]`, `[ERROR]` prefixes

## Dependencies

- **XLKit**: Main library for Excel generation
- **Foundation**: File system operations and utilities

## Future Enhancements

- [ ] Image embedding tests
- [ ] CSV import/export tests
- [ ] Cell formatting tests
- [ ] Chart generation tests
- [ ] Conditional formatting tests
- [ ] Multi-sheet workbook tests
- [ ] Performance benchmarking tests 