
def rgb_to_ansi(r, g, b):
    """
    Convert RGB values to the closest ANSI color code.
    """
    # Calculate the closest 8-bit color code
    if r == g == b:
        if r < 8:
            return 16
        if r > 248:
            return 231
        return round(((r - 8) / 247) * 24) + 232

    return 16 + (36 * round(r / 255 * 5)) + (6 * round(g / 255 * 5)) + round(b / 255 * 5)

# Define the RGB color
rgb_color = (49, 51, 158)

# Convert RGB to ANSI escape code for background
ansi_bg_code = rgb_to_ansi(*rgb_color)

PPT_BG_BLUE = ''

# Define ANSI escape code for white text
WHITE_TEXT = '\033[38;5;15m'

# Define ANSI escape codes for bold and color
BOLD = '\033[1m'
ITALIC = '\033[3m'
RESET = '\033[0m'
RED = '\033[31m'
GREEN = '\033[32m'
YELLOW = '\033[33m'
BLUE = '\033[34m'
GREY = '\033[37m'
PURPLE = '\033[35m'
CYAN = '\033[36m'
ORANGE = '\033[38;5;208m'
PINK = '\033[38;5;201m'
LIGHT_BLUE = '\033[38;5;123m'
# Cool ANSI escape codes
# BLINK is an ANSI escape code for blinking text
# Example usage: print(f"{BLINK}This text will blink{RESET}")
BLINK = '\033[5m'
UNDERLINE = '\033[4m'
REVERSE = '\033[7m'
# PPT_BLUE = f'\033[38;5;{ansi_code}m'

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
