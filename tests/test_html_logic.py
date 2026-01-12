import unittest
import html
import re

# Mocking pptx Run object for testing
class MockFont:
    def __init__(self, bold=None, italic=None, underline=None, strike=None, size_pt=None, color_rgb=None, color_theme=None):
        self.bold = bold
        self.italic = italic
        self.underline = underline
        # In python-pptx, strict strikethrough is .font.strike, or sometimes .strike on run depending on version.
        # We will assume .font.strike as per code trace.
        self.strike = strike
        self.size = None
        if size_pt:
            self.size = type('obj', (object,), {'pt': size_pt})

        self.color = type('obj', (object,), {'type': None, 'rgb': None, 'theme_color': None})
        if color_rgb:
            self.color.type = 1 # MSO_COLOR_TYPE.RGB
            self.color.rgb = color_rgb
        elif color_theme:
            self.color.type = 2 # MSO_COLOR_TYPE.THEME
            self.color.theme_color = color_theme

class MockRun:
    def __init__(self, text, font=None):
        self.text = text
        self.font = font or MockFont()

class HTMLConverter:
    @staticmethod
    def run_to_html(run):
        text = html.escape(run.text)
        if not text:
            return ""

        # Color and Size
        style_parts = []

        # Size
        if run.font.size and run.font.size.pt:
            style_parts.append(f"font-size:{int(run.font.size.pt)}pt")

        # Color
        if run.font.color.type == 1: # RGB
            style_parts.append(f"color:#{run.font.color.rgb}")
        elif run.font.color.type == 2: # Theme
            # We encode theme as a data attribute on the span, or strictly in style if we could map it.
            # But here we will use a data attribute for reconstruction.
            pass

        # Construct tags
        # We wrap inner to outer: Text -> Bold -> Italic -> Underline -> Strike -> Span (Color/Size)

        result = text

        if run.font.bold:
            result = f"<b>{result}</b>"
        if run.font.italic:
            result = f"<i>{result}</i>"
        if run.font.underline:
            result = f"<u>{result}</u>"
        if run.font.strike:
            result = f"<s>{result}</s>"

        # Span for style and attributes
        attrs = []
        if style_parts:
            attrs.append(f'style="{"; ".join(style_parts)}"')

        if run.font.color.type == 2:
            attrs.append(f'data-pptx-theme-color="{run.font.color.theme_color}"')

        if attrs:
            result = f"<span {' '.join(attrs)}>{result}</span>"

        return result

    @staticmethod
    def html_to_runs(html_text):
        # We need a robust parser. Since we generated it, it's mostly flat or simple nesting.
        # But LLM might mess it up.
        # Let's use a regex-based tokenizer for simplicity if structure is flat-ish,
        # OR use html.parser. HTMLParser is better.

        from html.parser import HTMLParser

        class RunParser(HTMLParser):
            def __init__(self):
                super().__init__()
                self.runs = []
                self.current_style = {
                    "bold": False, "italic": False, "underline": False, "strike": False,
                    "font_size": None, "color_rgb": None, "theme_color": None
                }
                # Stack to track nested styles
                self.style_stack = []

            def handle_starttag(self, tag, attrs):
                # Push current style to stack
                self.style_stack.append(self.current_style.copy())

                attrs_dict = dict(attrs)

                if tag == "b" or tag == "strong":
                    self.current_style["bold"] = True
                elif tag == "i" or tag == "em":
                    self.current_style["italic"] = True
                elif tag == "u":
                    self.current_style["underline"] = True
                elif tag == "s" or tag == "strike" or tag == "del":
                    self.current_style["strike"] = True
                elif tag == "span" or tag == "font":
                    # Parse style
                    style_str = attrs_dict.get("style", "")
                    # Simple style parser
                    styles = [s.strip().split(":") for s in style_str.split(";") if ":" in s]
                    for k, v in styles:
                        k = k.strip().lower()
                        v = v.strip().lower()
                        if k == "font-size":
                            if "pt" in v:
                                try:
                                    self.current_style["font_size"] = float(v.replace("pt", ""))
                                except:
                                    pass
                        elif k == "color":
                            if v.startswith("#"):
                                self.current_style["color_rgb"] = v.replace("#", "").upper()

                    # Theme color data attribute
                    if "data-pptx-theme-color" in attrs_dict:
                        self.current_style["theme_color"] = attrs_dict["data-pptx-theme-color"]

                    # Handle <font color> tag just in case
                    if tag == "font":
                        if "color" in attrs_dict:
                            c = attrs_dict["color"]
                            if c.startswith("#"):
                                self.current_style["color_rgb"] = c.replace("#", "").upper()

            def handle_endtag(self, tag):
                if self.style_stack:
                    self.current_style = self.style_stack.pop()

            def handle_data(self, data):
                if not data:
                    return
                # Create a run spec
                self.runs.append({
                    "text": data, # unescape handled by HTMLParser? No, usually handle_data receives unescaped.
                                  # Wait, HTMLParser.handle_data receives the raw text (already decoded entites usually).
                                  # Let's verify.
                    "style": self.current_style.copy()
                })

        parser = RunParser()
        parser.feed(html_text)
        return parser.runs

class TestHTMLLogic(unittest.TestCase):
    def test_run_to_html_simple(self):
        run = MockRun("Hello", MockFont(bold=True))
        html_out = HTMLConverter.run_to_html(run)
        self.assertEqual(html_out, "<b>Hello</b>")

    def test_run_to_html_complex(self):
        run = MockRun("World", MockFont(italic=True, size_pt=12, color_rgb="FF0000"))
        html_out = HTMLConverter.run_to_html(run)
        # Order of attributes might vary if we used dict, but list is deterministic
        # Expected: <span style="font-size:12pt; color:#FF0000"><i>World</i></span>
        self.assertIn("font-size:12pt", html_out)
        self.assertIn("color:#FF0000", html_out)
        self.assertIn("<i>World</i>", html_out)
        self.assertTrue(html_out.startswith("<span"))

    def test_html_to_runs_simple(self):
        html_in = "<b>Hello</b>"
        runs = HTMLConverter.html_to_runs(html_in)
        self.assertEqual(len(runs), 1)
        self.assertEqual(runs[0]["text"], "Hello")
        self.assertTrue(runs[0]["style"]["bold"])
        self.assertFalse(runs[0]["style"]["italic"])

    def test_html_to_runs_nested(self):
        html_in = '<span style="color:#00FF00">Prefix <b>Bold</b> Suffix</span>'
        runs = HTMLConverter.html_to_runs(html_in)
        # Should be 3 runs
        # 1. Prefix (green)
        # 2. Bold (green, bold)
        # 3. Suffix (green)
        self.assertEqual(len(runs), 3)
        self.assertEqual(runs[0]["text"], "Prefix ")
        self.assertEqual(runs[0]["style"]["color_rgb"], "00FF00")

        self.assertEqual(runs[1]["text"], "Bold")
        self.assertTrue(runs[1]["style"]["bold"])
        self.assertEqual(runs[1]["style"]["color_rgb"], "00FF00")

        self.assertEqual(runs[2]["text"], " Suffix")
        self.assertFalse(runs[2]["style"]["bold"])
        self.assertEqual(runs[2]["style"]["color_rgb"], "00FF00")

    def test_broken_html_resilience(self):
        # LLM might forget to close tags
        html_in = "<b>Unclosed bold"
        runs = HTMLConverter.html_to_runs(html_in)
        self.assertEqual(len(runs), 1)
        self.assertEqual(runs[0]["text"], "Unclosed bold")
        self.assertTrue(runs[0]["style"]["bold"])

if __name__ == "__main__":
    unittest.main()
