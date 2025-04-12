# WordDocumentMediaExtractor
<p align="center">
  <img src="Distribution/WordExtractorIcon.ico" alt="Icon" width="200"/>
</p>

A collection of simple Python scripts useful for extracting images, GIFs, videos, and other media from `.doc` and `.docx` documents.

## How to Use

1. Place your documents in the `Source` folder.
2. Run the `WordExtractorRunner.py` script.

Alternatively, you can generate an executable `.exe` file by navigating to the `Distribution` folder and running the `.bat` script. This will create an executable file that you can use anywhere. Just make sure to keep your documents in the same folder as the `.exe` file.

## Output

For each document processed, a new folder will be created containing:
- The parsed text in `.json` format.
- The extracted media files.