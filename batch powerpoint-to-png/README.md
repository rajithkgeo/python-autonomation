# Batch PowerPoint to PNG Converter

A Python automation utility that batch-processes PowerPoint presentations
and exports every slide as a high-resolution PNG image.

## Overview

When a large number of PowerPoint presentations need to be converted
into images, manually exporting each presentation and each slide is
time-consuming.

This tool automates the process using Microsoft PowerPoint's COM
interface.

## Workflow

```text
PowerPoint files
(.ppt / .pptx)
        │
        ▼
Batch discovery
        │
        ▼
Open PowerPoint automatically
        │
        ▼
Read slide dimensions
        │
        ▼
Calculate target resolution
        │
        ▼
Export slides as PNG
        │
        ▼
Organized output folders
```
## Features
- Batch process multiple .ppt and .pptx files
- Automatically ignore temporary PowerPoint files
- Export every slide as PNG
- Preserve the slide aspect ratio
- Calculate output dimensions from target DPI
- Create a separate output folder for each presentation
- Process presentations without opening visible PowerPoint windows
- Close PowerPoint automatically after processing

## Requirements
### Operating system
- Windows
### Software
- Microsoft PowerPoint desktop
- Python 3.x
### Python package
```text
pip install pywin32
```
## Running Instructions
After running the code follow the console prompts to enter your input directory, output directory, 
and target DPI (or press Enter to accept the defaults)
## Output structure

If the input directory contains:
```text
Input/
├── Presentation_A.pptx
├── Presentation_B.pptx
└── Presentation_C.pptx
```
The output will be:
```text
Output/
├── Presentation_A/
│   ├── Slide_01.png
│   ├── Slide_02.png
│   └── ...
│
├── Presentation_B/
│   ├── Slide_01.png
│   ├── Slide_02.png
│   └── ...
│
└── Presentation_C/
    ├── Slide_01.png
    ├── Slide_02.png
    └── ...
```
## Resolution

The output resolution is calculated from the PowerPoint slide dimensions
and the selected DPI.

For example:
```text
dpi = 300
```
The script reads the slide width and height in points and converts them
to the corresponding pixel dimensions before exporting the slide.

## Usage
Each image can then be used for:

- posters
- digital displays
- video/slideshow creation
- presentations
- web publishing
- image archives
## Limitations
- Windows only
- Requires Microsoft PowerPoint desktop
- The current implementation uses PowerPoint's own rendering engine
- Very large batches may take significant processing time
## Author

Rajith K

Python Automation | GIS | Remote Sensing | Geospatial Technology
