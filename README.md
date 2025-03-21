# Docx to PDF Converter

A lightweight npm package designed to convert Word documents (DOCX) to PDF format. In addition to conversion, it now includes support for extracting images from DOCX files. This package provides a simple and efficient solution for handling DOCX to PDF conversion in a cross-platform manner.

## Features
- Cross-platform compatibility: Supports Windows, macOS, and Linux.
- Easy-to-use API: Convert DOCX files to PDF with just a few lines of code.
- Batch conversion: Convert entire directories of DOCX files to PDF.
- Image extraction: Extract images embedded in DOCX files and save them to a specified directory.

## Installation
Install the package via npm:

```bash
npm i docx2pdf-converter
```
Install the package via yarn:

```bash
yarn add docx2pdf-converter
```

## Usage
### Convert a Single DOCX to PDF
This line of code converts a single DOCX file into a PDF file:

```javascript
const DOCX2PDFConverter = require('docx2pdf-converter');
const path = require('path');

const inputPath = path.join(__dirname, "./report.docx");
DOCX2PDFConverter.convert(inputPath, path.join(__dirname, "output.pdf"))
  .then((res) => {
    console.log(res);
  })
  .catch(error=> {
    console.error(error)
  });
```

### Convert an Entire Directory of DOCX Files to PDF
This code converts all DOCX files within a directory to PDFs and saves them to another directory:

```javascript
const fs = require('fs');
const path = require('path');
const { convert, resolvePaths } = require('docx2pdf-converter');

async function convertDirectory(inputDir, outputDir) {
  const files = fs.readdirSync(inputDir);

  for(const file of files) {
    if (file.endsWith('.docx')) {
      const inputPath = path.join(inputDir, file);
      const { output } = resolvePaths(inputPath, outputDir);
      try {
        const result = await convert(inputPath, output);
        console.log(`Converted: ${file}`, result);
      } catch(err) {
        console.error(err);
      }
    }
  }
}

/*
Assume both directories (input and output) are in the same folder.
If not, you can provide absolute paths to the folders.
*/
const inputDirectory = path.join(__dirname, './inputdir');
const outputDirectory = path.join(__dirname, './outputdir');

convertDirectory(inputDirectory, outputDirectory)
  .then(() => {
    console.log("convert success !");
  })
  .catch(console.error);
```

### Extract Images from a DOCX File
The `extractImages` function allows you to extract all images from a DOCX file and save them to a specified directory:

```javascript
const { extractImages } = require('docx2pdf-converter');
const path = require('path');
const fs = require('fs');

const inputPath = path.join(__dirname, './report.docx');  // Path to your DOCX file
const outputDir = path.join(__dirname, './extracted-images');  // Directory where images will be saved

if(!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir);
}

extractImages(inputPath, outputDir)
  .then((result) => {
    console.log("convert success !", result);
  })
  .catch(console.error);;
```

This will extract any images embedded in the DOCX file and save them in the `extracted-images` directory.

## API
### `convert(inputPath, outputPath, keepActive = false)`
**Description**: Converts a DOCX file to PDF.

**Parameters**:
- `inputPath` (*string*): Path to the input DOCX file.
- `outputPath` (*string*): Path to the output PDF file.
- `keepActive` (*boolean, optional*): Flag to keep the application active (platform-dependent).

**Returns**: string. stdout.

### `extractImages(inputPath, outputDir)`
**Description**: Extracts images from a DOCX file and saves them to the specified directory.

**Parameters**:
- `inputPath` (*string*): Path to the input DOCX file.
- `outputDir` (*string*): Directory where the extracted images will be saved.

**Returns**: boolean. It extracts the images and logs the status.

### `resolvePaths(inputPath, outputPath)`
**Description**: Resolves and validates input and output paths, ensuring they are correct and handle both single files and directories.

**Parameters**:
- `inputPath` (*string*): Path to the input DOCX file or directory.
- `outputPath` (*string, optional*): Path to the output directory or file.

**Returns**: An object containing resolved input and output paths, and a batch flag indicating whether it's batch processing.

## Platform-Specific Support
- **Windows**: Uses PowerShell scripts (`convert.ps1` and `convertTodocx.ps1`) for converting DOCX to PDF and vice versa.
- **macOS**: Uses a shell script (`convert.sh`) for converting DOCX to PDF.
- **Linux**: Uses unoconv to convert DOCX to PDF.

The `convert` function automatically detects the platform and uses the appropriate method for conversion.

## Conclusion
This package provides a flexible, cross-platform solution for converting DOCX files to PDF and extracting images from DOCX files. Whether you're dealing with individual files or batch processing, the API makes it easy to automate the conversion and image extraction processes.