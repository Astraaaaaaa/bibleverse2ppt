# bibleverse2ppt

<!-- ![ASCII Art](carbon.png) -->

```
    __    _ __    __                             ___               __ 
   / /_  (_) /_  / /__ _   _____  _____________ |__ \ ____  ____  / /_
  / __ \/ / __ \/ / _ \ | / / _ \/ ___/ ___/ _ \__/ // __ \/ __ \/ __/
 / /_/ / / /_/ / /  __/ |/ /  __/ /  (__  )  __/ __// /_/ / /_/ / /_  
/_.___/_/_.___/_/\___/|___/\___/_/  /____/\___/____/ .___/ .___/\__/  
                                                  /_/   /_/           
```

# Bible Verse PowerPoint Generator from Text File

This project provides a script to generate a PowerPoint presentation from a text file containing Bible verses. The script reads the text file, processes the content, and creates a PowerPoint presentation with customizable background colors, font colors, and font sizes.

## Features

- Generate PowerPoint slides from a text file containing Bible verses.
- Customize background color, font color, and font size.
- Optionally add a background image with adjustable transparency.
- Supports UTF-8 encoding for text files.
- ANSI escape codes for colorful console output.

## Requirements

![Python](https://img.shields.io/badge/Python-3.x-blue.svg)
![python-pptx](https://img.shields.io/badge/python--pptx-0.6.21-green.svg)
![Pillow](https://img.shields.io/badge/Pillow-8.2.0-yellow.svg)
![lxml](https://img.shields.io/badge/lxml-4.6.3-red.svg)
![Selenium](https://img.shields.io/badge/Selenium-3.141.0-orange.svg)
![BeautifulSoup](https://img.shields.io/badge/BeautifulSoup-4.9.3-purple.svg)

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/Astraaaaaaa/bibleverse2ppt.git
    cd bibleverse2ppt
    ```

2. Install the required Python libraries:
    ```sh
    pip install -r requirements.txt
    ```

## Usage

### Command Line

To generate a PowerPoint presentation from a text file containing Bible verses, use the following command:

```sh
python get_bible_verse_from_website.py \ 
    --input query_bible.txt \ 
    --output verse.pptx \ 
    --bg-color red \ 
    --font-color white \ 
    --font-size 48 \ 
```

### Arguments

- `--input`: Path to the input text file containing Bible verses (default: `query_bible.txt`).
- `--output`: Path to the output PowerPoint file (default: `verse.pptx`).
- `--bg-color`: Background color (default: `default`).
- `--font-color`: Font color (default: `white`).
- `--font-size`: Font size (default: `48`).

### Sample Input File

Create a text file named `query_bible.txt` with the following format:

```
賽1:8-12

詩150

詩121:2-6

歌羅西書1:2-5

路加福音1:8-12
```

- Each line represents a query for Bible verses.
- The format is `BookNameChapter:VerseRange`.
- The script will fetch the specified verses and generate slides accordingly.

## Maintainer

Astra Lee <astralee95@gmail.com>

If you have any further questions or need additional modifications, feel free to ask!
