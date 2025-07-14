# Test-Data

This folder contains sample datasets used by the XLKitTestRunner for testing and demonstration purposes.

## Datasets

### Embed-Test

The `Embed-Test/` folder contains a sample dataset for testing image embedding functionality. This dataset was created using [MarkerData](https://github.com/TheAcharya/MarkerData).

- Embed-Test.csv - CSV file with 22 columns and 4 data rows
- Embed-Test_00-00-08-06.png - Sample PNG image (1920x1080)
- Embed-Test_00-00-22-07.png - Sample PNG image (1920x1080)
- Embed-Test_00-00-50-08.png - Sample PNG image (1920x1080)
- Embed-Test_00-01-09-10.png - Sample PNG image (1920x1080)

This dataset is used by the `embed` and `no-embeds` test types to demonstrate:
- CSV import functionality
- Image embedding with perfect aspect ratio preservation
- Column width optimization
- Cell formatting and styling

## Usage

The test runner automatically uses these datasets when running tests:

```bash
swift run XLKitTestRunner embed      # Uses Embed-Test dataset with images
swift run XLKitTestRunner no-embeds  # Uses Embed-Test dataset without images
```

## Data Format

The CSV file contains sample video editing data with columns for:
- Marker information (ID, name, type, status)
- Clip details (name, duration, keywords)
- Project metadata (reel, scene, take)
- Image filenames for embedding

All images are in PNG format with 1920x1080 resolution for testing aspect ratio preservation.

## Data Source

The sample data was generated using [MarkerData](https://github.com/TheAcharya/MarkerData), which allows users to extract, convert & create databases from Final Cut Pro's Marker metadata effortlessly. MarkerData provides batch extraction and rendering of stills based on each Marker's timecode, making it perfect for creating test datasets for Excel generation tools like XLKit. 