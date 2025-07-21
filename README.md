# Export Excel

A flexible and extensible Java library for exporting data to Excel files (`.xlsx`), supporting custom formatting, templates, and large data sets. Built on Apache POI SXSSF for performance and scalability.

## Features

- Export Java data to Excel with customizable templates
- Supports large data sets via streaming (SXSSF)
- Extensible: add custom cell data and style handlers
- Multi-sheet export with parallel processing
- Easy integration with any Java application

## Getting Started

### Prerequisites

- Java 8+
- Maven

### Installation

Add the module to your Maven project:

```xml
<dependency>
    <groupId>vn.com.phat.excel</groupId>
    <artifactId>export-excel</artifactId>
    <version>1.0.0-RELEASE</version>
</dependency>
```

### Basic Usage

```java
ExcelExporter exporter = new ExcelExporterDefaultImpl();
ExcelExporterParam param = ExcelExporterParamDefault.builder()
    .sheets(sheets) // List<ExcelExporterSheetParam>
    .datePattern("yyyy-MM-dd")
    .outputName("report.xlsx")
    .templatePath("/path/to/template.xlsx")
    .build();

try (OutputStream out = new FileOutputStream("output.xlsx")) {
    exporter.doExport(out, param);
}
```

## Extending the Module

- Implement new `CellDataTypeHandler` or `CellStyleHandler` for custom cell types or styles.
- See the `cellhandler` and `cellstyle` packages for examples.

## Project Structure

- `core/excel/ExcelExporterDefaultImpl.java` – Main export logic
- `core/excel/param/` – Parameter objects for configuration
- `core/excel/cellhandler/` – Cell data handlers
- `core/excel/cellstyle/` – Cell style handlers
- `core/excel/util/` – Utilities for data and column extraction

## Error Handling & Troubleshooting

- Ensure the template file exists and is a valid `.xlsx` file.
- All data should be mapped to Java objects before export.
- For large exports, monitor memory usage and adjust JVM settings if needed.

## License

This project is licensed under the [Apache License 2.0](LICENSE).
