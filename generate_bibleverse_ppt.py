from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup
import time

from webdriver_manager.chrome import ChromeDriverManager

from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml
from pptx.dml.color import RGBColor
from pptx.oxml.ns import nsdecls

from lxml import etree
from PIL import Image

# Set up Chrome options
options = Options()
options.add_argument("--disable-notifications")  # Disable notifications
options.add_argument("blink-settings=imagesEnabled=false")  # Disable images
options.add_argument("--headless")  # Run in headless mode
options.add_argument("--ignore-certificate-errors")
options.add_argument("--allow-running-insecure-content")

# Define Old Testament books with their initials, full names, and number of chapters

# Define Old Testament books with their initials, full names, number of chapters, and verse limits
old_testament_books = [
    ("創", "創世紀", 50, [31, 25, 24, 26, 33, 22, 24, 22, 29, 32, 30, 20, 18, 24, 21, 16, 27, 30, 37, 18, 34, 31, 20, 18, 34, 22, 36, 30, 35, 43, 55, 30, 31, 29, 30, 31, 29, 30, 31, 30, 31, 30, 31, 30, 31, 30, 31, 30, 31, 26]),  # Genesis
    ("出", "出埃及記", 40, [22, 25, 22, 31, 23, 30, 29, 28, 35, 29, 10, 51, 22, 31, 27, 36, 16, 27, 25, 26, 36, 31, 27, 18, 40, 37, 29, 43, 46, 38, 35, 38, 35, 29, 35, 29, 30, 31, 29, 38]),  # Exodus
    ("利", "利未記", 27, [17, 16, 17, 35, 19, 30, 38, 36, 24, 20, 47, 8, 59, 57, 33, 34, 34, 30, 37, 27, 24, 33, 44, 23, 55, 46, 34]),  # Leviticus
    ("民", "民數記", 36, [54, 34, 51, 49, 31, 27, 41, 26, 23, 36, 35, 16, 33, 34, 41, 30, 28, 32, 37, 29, 35, 34, 30, 25, 18, 23, 23, 31, 30, 31, 30, 36, 34, 29, 30, 34]),  # Numbers
    ("申", "申命記", 34, [46, 37, 29, 49, 33, 25, 26, 20, 29, 22, 32, 32, 25, 31, 29, 30, 26, 22, 25, 20, 23, 30, 29, 25, 29, 30, 26, 24, 29, 30, 31, 30, 29, 30]),  # Deuteronomy
    ("書", "約書亞記", 24, [18, 24, 17, 24, 15, 27, 26, 35, 27, 43, 15, 24, 33, 15, 20, 10, 18, 28, 24, 27, 34, 20, 30, 31]),  # Joshua
    ("士", "士師記", 21, [36, 23, 31, 24, 31, 40, 25, 35, 57, 18, 40, 15, 25, 20, 31, 31, 30, 31, 30, 30, 25]),  # Judges
    ("得", "路得記", 4, [22, 23, 18, 22]),  # Ruth
    ("撒上", "撒母耳記上", 31, [28, 36, 21, 22, 27, 21, 29, 22, 27, 25, 15, 25, 23, 22, 35, 23, 27, 30, 24, 42, 15, 25, 29, 22, 24, 25, 12, 25, 24, 31, 13]),  # 1 Samuel
    ("撒下", "撒母耳記下", 24, [27, 32, 39, 22, 25, 23, 29, 18, 25, 19, 27, 31, 39, 33, 24, 23, 29, 27, 43, 26, 17, 25, 39, 25]),  # 2 Samuel
    ("王上", "列王紀上", 22, [51, 46, 28, 34, 32, 38, 51, 66, 29, 29, 43, 32, 34, 31, 30, 34, 24, 38, 37, 43, 29, 53]),  # 1 Kings
    ("王下", "列王紀下", 25, [27, 25, 27, 37, 27, 33, 20, 29, 28, 29, 37, 21, 25, 30, 30, 20, 25, 30, 30, 30, 25, 30, 30, 30, 30]),  # 2 Kings
    ("代上", "歷代志上", 29, [54, 41, 24, 43, 29, 31, 40, 40, 44, 44, 47, 22, 25, 22, 29, 27, 27, 32, 36, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]),  # 1 Chronicles
    ("代下", "歷代志下", 36, [15, 35, 20, 22, 26, 23, 22, 27, 29, 36, 22, 23, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]),  # 2 Chronicles
    ("拉", "以斯拉記", 10, [11, 70, 13, 24, 17, 22, 28, 36, 15, 44]),  # Ezra
    ("尼", "尼希米記", 13, [11, 20, 38, 23, 19, 16, 73, 18, 38, 36, 36, 47, 31]),  # Nehemiah
    ("斯", "以斯帖記", 10, [22, 23, 15, 17, 14, 13, 10, 17, 15, 3]),  # Esther
    ("伯", "約伯記", 42, [22, 18, 25, 21, 27, 30, 21, 22, 35, 22, 20, 25, 28, 22, 31, 22, 20, 21, 29, 29, 34, 30, 25, 22, 30, 23, 25, 22, 30, 31, 22, 30, 22, 30, 22, 30, 22, 30, 22, 30, 22, 30]),  # Job
    ("詩", "詩篇", 150, [41, 50, 31, 68, 36, 36, 83, 51, 66, 72, 89, 110, 13, 31, 71, 13, 24, 17, 104, 17, 23, 19, 15, 19, 11, 23, 13, 28, 23, 19, 37, 36, 12, 29, 17, 18, 51, 29, 80, 55, 25, 31, 57, 79, 45, 49, 41, 42, 74, 23, 13, 20, 65, 36, 36, 42, 84, 45, 51, 39, 49, 31, 27, 60, 40, 43, 13, 31, 7, 10, 33, 37, 71, 18, 53, 100, 16, 24, 20, 23, 11, 13, 21, 72, 13, 20, 17, 8, 19, 13, 14, 17, 7, 19, 53, 17, 16, 16, 5, 23, 11, 13, 12, 9, 9, 5, 8, 29, 22, 35, 45, 48, 43, 14, 31, 7, 10, 10, 9, 26, 18, 19, 2, 29, 176, 7, 8, 9, 4, 8, 5, 6, 5, 6, 8, 8, 3, 18, 3, 3, 21, 26, 9, 8, 24, 14, 10, 8, 12, 15, 21, 10, 20, 14, 9, 6]),
    ("箴", "箴言", 31, [22, 24, 35, 27, 23, 35, 27, 36, 18, 32, 31, 28, 25, 35, 29, 33, 28, 24, 29, 30, 31, 30, 28, 31, 28, 27, 29, 28, 27, 30, 31]),  # Proverbs
    ("傳", "傳道書", 12, [18, 26, 22, 16, 20, 12, 29, 17, 18, 20, 10, 14]),  # Ecclesiastes
    ("歌", "雅歌", 8, [17, 13, 11, 16, 16, 13, 13, 14]),  # Song of Solomon
    ("賽", "以賽亞書", 66, [31, 22, 26, 6, 30, 13, 25, 22, 21, 34, 16, 6, 12, 32, 9, 14, 10, 14, 25, 6, 17, 25, 24, 23, 12, 21, 13, 22, 24, 31, 9, 20, 21, 22, 24, 31, 20, 22, 29, 31, 29, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31]),  # Isaiah
    ("耶", "耶利米書", 52, [19, 37, 25, 30, 31, 30, 34, 23, 24, 25, 23, 17, 27, 23, 21, 21, 27, 23, 25, 18, 14, 30, 30, 29, 38, 24, 22, 29, 32, 24, 26, 22, 30, 24, 22, 30, 22, 30, 22, 30, 22, 30, 22, 30, 22, 30, 22, 30, 22, 30, 22, 30]),  # Jeremiah
    ("哀", "耶利米哀歌", 5, [22, 22, 66, 22, 22]),  # Lamentations
    ("結", "以西結書", 48, [28, 10, 27, 17, 17, 14, 27, 18, 27, 22, 25, 28, 23, 23, 25, 24, 27, 24, 23, 24, 25, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]),  # Ezekiel
    ("但", "但以理書", 12, [21, 49, 30, 37, 31, 28, 28, 27, 27, 21, 45, 13]),  # Daniel
    ("何", "何西阿書", 14, [11, 23, 5, 19, 15, 11, 16, 14, 17, 15, 12, 14, 16, 9]),  # Hosea
    ("摩", "摩西亞書", 9, [15, 16, 15, 13, 15, 14, 15, 14, 15]),  # Amos
    ("俄", "俄巴底亞書", 1, [21]),  # Obadiah
    ("拿", "拿鴻書", 3, [15, 13, 19]),  # Nahum
    ("彌", "彌迦書", 7, [16, 13, 12, 13, 15, 16, 20]),  # Micah
    ("鴻", "哈巴谷書", 3, [17, 20, 19]),  # Habakkuk
    ("哈", "哈該書", 2, [15, 23]),  # Haggai
    ("亞", "撒迦利亞書", 14, [21, 13, 10, 14, 11, 15, 14, 23, 17, 12, 17, 14, 9, 21]),  # Zechariah
    ("瑪", "瑪拉基書", 4, [14, 17, 18, 6])  # Malachi
]

new_testament_books = [
    ("太", "馬太福音", 28, [25, 23, 17, 25, 48, 34, 29, 34, 38, 42, 30, 50, 58, 36, 39, 28, 27, 35, 30, 34, 46, 46, 39, 51, 46, 75, 66, 20]),
    ("可", "馬可福音", 16, [45, 28, 35, 41, 43, 56, 37, 38, 50, 52, 33, 44, 37, 72, 47, 20]),
    ("路", "路加福音", 24, [80, 52, 38, 44, 39, 49, 50, 56, 62, 42, 54, 59, 35, 35, 32, 31, 37, 43, 48, 47, 38, 71, 56, 53]),
    ("約", "約翰福音", 21, [51, 25, 36, 54, 47, 71, 53, 59, 41, 42, 57, 50, 38, 31, 27, 33, 26, 40, 42, 31, 25]),
    ("徒", "使徒行傳", 28, [26, 47, 26, 37, 42, 15, 60, 40, 43, 48, 30, 25, 52, 28, 41, 40, 34, 28, 40, 38, 40, 30, 35, 27, 27, 32, 44, 31]),
    ("羅", "羅馬書", 16, [32, 29, 31, 25, 21, 23, 25, 39, 33, 21, 36, 21, 14, 23, 33, 27]),
    ("林前", "哥林多前書", 16, [31, 16, 23, 21, 13, 20, 40, 13, 27, 33, 34, 31, 13, 40, 58, 24]),
    ("林後", "哥林多後書", 13, [24, 17, 18, 18, 21, 18, 16, 24, 15, 18, 33, 21, 14]),
    ("加", "加拉太書", 6, [24, 21, 29, 31, 26, 18]),
    ("弗", "以弗所書", 6, [23, 22, 21, 32, 33, 24]),
    ("腓", "腓立比書", 4, [30, 30, 21, 23]),
    ("西", "歌羅西書", 4, [29, 23, 25, 18]),
    ("帖前", "帖撒羅尼迦前書", 5, [10, 20, 13, 18, 28]),
    ("帖後", "帖撒羅尼迦後書", 3, [12, 17, 18]),
    ("提前", "提摩太前書", 6, [20, 15, 16, 16, 25, 21]),
    ("提後", "提摩太後書", 4, [18, 26, 17, 22]),
    ("多", "提多書", 3, [16, 15, 15]),
    ("門", "腓利門書", 1, [25]),
    ("來", "希伯來書", 13, [14, 18, 19, 16, 14, 20, 28, 13, 28, 39, 40, 29, 25]),
    ("雅", "雅各書", 5, [27, 26, 18, 17, 20]),
    ("彼前", "彼得前書", 5, [25, 25, 22, 19, 14]),
    ("彼後", "彼得後書", 3, [21, 22, 18]),
    ("約一", "約翰一書", 5, [10, 29, 24, 21, 21]),
    ("約二", "約習二書", 1, [13]),
    ("約三", "約習三書", 1, [14]),
    ("猶", "猶大書", 1, [25]),
    ("啟", "啟示錄", 22, [20, 29, 22, 11, 14, 17, 17, 13, 21, 11, 19, 17, 18, 20, 8, 21, 18, 24, 21, 15, 27, 21]),
]

# Function to convert numeric values to Chinese traditional words
def num_to_chinese(num):
    chinese_nums = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    chinese_tens = ["", "十", "二十", "三十", "四十", "五十", "六十", "七十", "八十", "九十"]
    
    if num < 10:
        return chinese_nums[num]
    elif num < 20:
        return "十" + chinese_nums[num % 10]
    else:
        return chinese_tens[num // 10] + chinese_nums[num % 10]

# Initialize the WebDriver with the options
browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

browser.get("https://springbible.fhl.net/Bible2/cgic201/read100.html")

result = []
# Read book name, chapter number, and verse range from query_bible.txt
with open("query_bible.txt", "r", encoding="utf-8") as file:
    for line in file:
        book_chap_verse = line.strip().split(':')
        book_name = ''.join(filter(str.isalpha, book_chap_verse[0]))
        
        # Check if there is a valid chapter number
        chap_str = ''.join(filter(str.isdigit, book_chap_verse[0]))
        chap_num = int(chap_str) if chap_str else None
        
        verse_range = book_chap_verse[1].strip()
        if '-' in verse_range:
            verse_range = [int(x) for x in verse_range.split('-')]
        else:
            verse_range = [int(verse_range), int(verse_range)]

        # Determine if the book is Old Testament or New Testament
        form_index = None
        chapter_limit = None

        # Check in Old Testament books
        for initial, full, chapters, verses in old_testament_books:
            if book_name in (initial, full):
                book_name = initial
                form_index = 1  # Use form[1] for Old Testament
                chapter_limit = chapters  # Get the number of chapters
                break

        # Check in New Testament books if not found in Old Testament
        if form_index is None:
            for initial, full, chapters, verses in new_testament_books:
                if book_name in (initial, full):
                    book_name = initial
                    form_index = 2  # Use form[2] for New Testament
                    chapter_limit = chapters  # Get the number of chapters

        # Handle the case where the book is not found
        if form_index is None:
            print(f"Book '{book_name}' is not found in either the Old Testament or New Testament.")
        else:
            # Check if the chapter number is within the limits
            if 1 <= chap_num <= chapter_limit:

                if form_index == 1:
                    if verse_range[0] <= len(old_testament_books[form_index - 1][3]) and verse_range[1] <= len(old_testament_books[form_index - 1][3]):
                        pass
                elif form_index == 2:
                    if verse_range[0] <= len(new_testament_books[form_index - 1][3]) and verse_range[1] <= len(new_testament_books[form_index - 1][3]):
                        pass
                    
                # Locate the specific form based on the determined index
                form = browser.find_element(By.XPATH, f"//form[{form_index}]")  # Adjust the index as needed

                # Find the select element by name within the specific form
                select_book = Select(form.find_element(By.NAME, "na"))
                select_book.select_by_value(book_name)

                # Find the select element by name
                select_chap = Select(form.find_element(By.NAME, "chap"))
                select_chap.select_by_value(f"{chap_num:03}")  # Format chap_num as a three-digit string

                # Find the submit button by name and click it
                submit_button = form.find_element(By.NAME, "submit1")
                submit_button.click()  # Click the submit button

                # Get the page source
                page_source = browser.page_source

                # Use BeautifulSoup to parse the HTML
                soup = BeautifulSoup(page_source, "html.parser")

                # Find all <ol> tags
                ol_elements = soup.find_all('ol')

                if form_index == 1:
                    full_book_name = next((full for initial, full, _, _ in old_testament_books if initial == book_name), "Unknown")
                elif form_index == 2:
                    full_book_name = next((full for initial, full, _, _ in new_testament_books if initial == book_name), "Unknown")
                
                chap_num_chinese = num_to_chinese(chap_num) + '章'
                verse = ''
                if verse_range[0] == verse_range[1]:
                    verse = str(verse_range[0])
                else:
                    verse = f"{verse_range[0]}-{verse_range[1]}"

                print(f"{full_book_name}{chap_num_chinese}{verse}節")
                
                title = f"{full_book_name}{chap_num_chinese}{verse}節"

                result.append(title)

                # Extract text from each <ol> and join them, filtering out empty lines
                ol_texts = []
                for i, ol in enumerate(ol_elements):
                    # Find all <li> items within the current <ol>
                    li_items = ol.find_all('li')
                    
                    # Number each <li> item
                    for j, li in enumerate(li_items):
                        text = li.get_text(strip=True)  # Get the text of the <li> and strip whitespace
                        if text and verse_range[0] <= j + 1 <= verse_range[1]:  # Only add text within the verse range
                            ol_texts.append(f"{j + 1}.{text.replace('神', ' 神')}")  # Format with index

                    # Join all the texts from <ol> elements into a single string
                    final_text = "\n".join(ol_texts) + '\n'  # Single newline between different items

                    result.append(final_text)
                    # Print the text from <ol> tags
                    print(final_text)

                    # Close the browser
                    browser.back()
                    
                    # print(f"Verse range {verse_range[0]}-{verse_range[1]} is within the verse limits for the book '{book_name}'.")
            else:
                print(f"Chapter number {chap_num} is outside the chapter limits for the book '{book_name}'.")

# Save the text to output.txt file
if result:
    # Convert the list elements to strings before writing to the file
    result_strings = [str(item) for item in result]

    # Write the list elements to the output.txt file
    with open("output.txt", "w", encoding="utf-8") as file:
        file.write("\n".join(result_strings))

def add_text_shadow(run):
    shadow_xml = """
    <a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:outerShdw blurRad="38100" dist="38100" dir="5400000" algn="ctr" rotWithShape="0">
            <a:srgbClr val="000000">
                <a:alpha val="50000"/>
            </a:srgbClr>
        </a:outerShdw>
    </a:effectLst>
    """
    shadow_element = parse_xml(shadow_xml)
    run_element = run._r
    run_properties = run_element.get_or_add_rPr()
    run_properties.append(shadow_element)

# Define a color mapping
COLOR_MAP = {
    "white": (255, 255, 255),
    "black": (0, 0, 0),
    "red": (255, 0, 0),
    "green": (0, 255, 0),
    "blue": (0, 0, 255),
    "yellow": (255, 255, 0),
    "purple": (128, 0, 128),
    "default": (49, 51, 158),
}

def set_background_color(slide, color):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*COLOR_MAP[color])

def generate_ppt_from_txt(txt_file, ppt_file, fontsize=30, fontcolor='white', bgcolor='default'):
    # Create a presentation object
    prs = Presentation()

    # Read the text file with UTF-8 encoding
    with open(txt_file, 'r', encoding='utf-8') as file:
        lines = file.read()

    if not lines:
        print("The text file is empty.")
        return

    # # Replace \x0b with \n
    # # content = content.replace('\x0b', '\n')

    # # Split by double newline and remove empty lines
    # slides_content = [slide.strip() for slide in lines if slide.split("\n\n")]

    for paragraph in lines.split("\n\n"):
        paragraph = paragraph.strip()
        for line in paragraph.split("\n")[1:]:
            slide = prs.slides.add_slide(prs.slide_layouts[1])  # Use the title and content layout
            title_shape = slide.shapes.title
            title_shape.text = paragraph.split("\n")[0]  # Set the title for the slide
            title_shape.text_frame.paragraphs[0].font.size = Pt(fontsize)
            title_shape.text_frame.paragraphs[0].font.underline = True
            title_shape.text_frame.paragraphs[0].font.bold = True
            title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(*COLOR_MAP[fontcolor])

            # Check for the first number in the sentence and use it for numbering
            first_number = line.split('.')[0]
            if first_number:
                pass

            # This line is part of the content
            text_frame = slide.placeholders[1].text_frame
            
            p = text_frame.add_paragraph()
            p.text = line #line.split('.')[1]  # Set the content line

            if text_frame.text and text_frame.text[0] == '\n':
                text_frame.text = text_frame.text[1:].strip()
                for i in range(len(text_frame.paragraphs)):
                    bu_num_id_val = etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}val")
                    bu_num_id_val.text = str(first_number)
                    bu_num_id = etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}buNumId")
                    bu_num_id.append(bu_num_id_val)
                    bu_num_pr = etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}buNumPr")
                    bu_num_pr.append(bu_num_id)
                    
                    # text_frame.paragraphs[i]._pPr.insert(0, bu_num_pr)
                    text_frame.paragraphs[i]._pPr.insert(0, etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}buNone"))
                    
                    text_frame.paragraphs[i].alignment = PP_ALIGN.JUSTIFY
                    text_frame.paragraphs[i].font.size = Pt(fontsize)
                    text_frame.paragraphs[i].level = 0
                    text_frame.paragraphs[0].font.bold = True
                    text_frame.paragraphs[i].font.color.rgb = RGBColor(*COLOR_MAP[fontcolor])
            
                    # # Insert the numbering property
                    # numPr = parse_xml(
                    #     f'<a:numPr {nsdecls("a")}>'
                    #     '<a:ilvl val="0"/>'  # Set the indentation level (0 for top-level)
                    #     '<a:startAt val="10"/>'  # Start numbering from 1
                    #     '</a:numPr>'
                    # )
                    # text_frame.paragraphs[i]._pPr.insert(0, numPr)


                    # text_frame.paragraphs[i]._pPr.insert(0, etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}buNone"))
                    # text_frame.paragraphs[i]._pPr.insert(0, etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}buNumPr"))
                    # text_frame.paragraphs[i]._pPr.insert(0, etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}buNumId"))

                    # # text_frame.paragraphs[i].font.bold = True
                    # # text_frame.paragraphs[i].font.color.rgb = RGBColor(*COLOR_MAP[fontcolor])
                    # # for run in text_frame.paragraphs[i].runs:
                    # #     add_text_shadow(run)  # Add shadow to each run in the paragraph
                    set_background_color(slide, 'default')

    try:
        prs.save(ppt_file)
        # print(f"{GREEN}{BOLD} >> Successfully saved PowerPoint presentation to {BLUE}{BOLD}{ppt_file_path}{RESET}")
    except Exception as e:
        print(f"Error: Failed to save PowerPoint presentation. {e}")
        # print(f"{RED}{BOLD}Error: Failed to save PowerPoint presentation. {e}{RESET}")

generate_ppt_from_txt("output.txt", "verse.pptx", fontsize=55)
