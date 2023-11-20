# Docx to PDF Converter

A simple and lightweight npm package for converting Word documents (docx) to PDF format using Node.js.

## Features

- Cross-platform compatibility: Supports both Windows and macOS.
- Easy-to-use API: Convert docx files to PDF with just a few lines of code.
- Batch conversion: Convert entire directories of docx files to PDF.

## Installation

Install the package via npm:

```bash
npm i docx2pdf-converter
```

## Usage

```javascript
const topdf = require('docx2pdf-converter')

const inputPath = './report.docx';

topdf.convert(inputPath,'output.pdf')
```


