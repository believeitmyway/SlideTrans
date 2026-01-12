import html
import re
from tqdm import tqdm
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from html.parser import HTMLParser

class HTMLRunParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.runs = []
        self.current_style = {
            "bold": False, "italic": False, "underline": False, "strike": False,
            "font_size": None, "color_rgb": None, "theme_color": None
        }
        self.style_stack = []

    def handle_starttag(self, tag, attrs):
        self.style_stack.append(self.current_style.copy())
        attrs_dict = dict(attrs)

        if tag in ["b", "strong"]:
            self.current_style["bold"] = True
        elif tag in ["i", "em"]:
            self.current_style["italic"] = True
        elif tag == "u":
            self.current_style["underline"] = True
        elif tag in ["s", "strike", "del"]:
            self.current_style["strike"] = True
        elif tag in ["span", "font"]:
            style_str = attrs_dict.get("style", "")
            styles = [s.strip().split(":") for s in style_str.split(";") if ":" in s]
            for k, v in styles:
                k = k.strip().lower()
                v = v.strip().lower()
                if k == "font-size":
                    if "pt" in v:
                        try:
                            self.current_style["font_size"] = float(v.replace("pt", ""))
                        except ValueError:
                            pass
                elif k == "color":
                    if v.startswith("#"):
                        self.current_style["color_rgb"] = v.replace("#", "").upper()

            if "data-pptx-theme-color" in attrs_dict:
                self.current_style["theme_color"] = attrs_dict["data-pptx-theme-color"]

            if tag == "font" and "color" in attrs_dict:
                c = attrs_dict["color"]
                if c.startswith("#"):
                    self.current_style["color_rgb"] = c.replace("#", "").upper()

    def handle_endtag(self, tag):
        if self.style_stack:
            self.current_style = self.style_stack.pop()

    def handle_data(self, data):
        if not data:
            return
        self.runs.append({
            "text": data,
            "style": self.current_style.copy()
        })

class PPTXProcessor:
    def __init__(self, filepath, translator):
        self.filepath = filepath
        self.translator = translator
        self.prs = Presentation(filepath)

    def process(self):
        """
        Iterates through all slides and shapes, processing text sequentially.
        """
        total_slides = len(self.prs.slides)
        print(f"Processing {total_slides} slides...")

        # We need to count tasks first for progress bar?
        # Or just iterate and be patient. Sequential processing is slower, so progress bar is good.
        # But counting requires full traversal. Let's do a generator or simple count.
        # Let's just process slide by slide and update tqdm per slide? Or per shape.
        # Per shape is better granularity.

        # To keep it simple and robust, let's collect all translate-able items first (just pointers),
        # then process them.
        all_tasks = []
        for slide in self.prs.slides:
            for shape in slide.shapes:
                self._collect_tasks(shape, all_tasks)

        print(f"Found {len(all_tasks)} text items to translate.")

        for task in tqdm(all_tasks, desc="Translating"):
            self._process_single_task(task)

    def _collect_tasks(self, shape, task_list, context="standard"):
        # Handle Groups (Recursive)
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for item in shape.shapes:
                self._collect_tasks(item, task_list, context="constrained")
            return

        # Handle Tables
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if not cell.text_frame:
                        continue
                    self._collect_text_frame_tasks(cell.text_frame, task_list, context="constrained")
            return

        # Handle Text Frames
        if shape.has_text_frame:
            self._collect_text_frame_tasks(shape.text_frame, task_list, context=context)

    def _collect_text_frame_tasks(self, text_frame, task_list, context="standard"):
        for paragraph in text_frame.paragraphs:
            if not paragraph.text.strip():
                continue

            task_list.append({
                "paragraph": paragraph,
                "context": context
            })

    def _process_single_task(self, task):
        paragraph = task["paragraph"]
        context = task["context"]

        # 1. Convert to HTML
        html_text = self._paragraph_to_html(paragraph)
        if not html_text:
            return

        # 2. Calculate limits
        raw_text_length = len("".join([r.text for r in paragraph.runs]))
        max_chars = self._calculate_max_chars(raw_text_length)

        # 3. Select Prompt
        if context == "constrained":
            prompt_template = self.translator.config.constrained_text_prompt
        else:
            prompt_template = self.translator.config.presentation_body_prompt

        # 4. Translate
        translated_html = self.translator.translate_text(html_text, max_chars, prompt_template)

        # 5. Reconstruct
        if translated_html:
            self._reconstruct_paragraph(paragraph, translated_html)

    def _paragraph_to_html(self, paragraph):
        """
        Converts paragraph runs to HTML string.
        """
        parts = []
        for run in paragraph.runs:
            parts.append(self._run_to_html(run))
        return "".join(parts)

    def _run_to_html(self, run):
        text = html.escape(run.text)
        if not text:
            return ""

        style_parts = []

        # Size
        if run.font.size and run.font.size.pt:
            style_parts.append(f"font-size:{int(run.font.size.pt)}pt")

        # Color
        try:
            if run.font.color.type == 1: # RGB
                style_parts.append(f"color:#{run.font.color.rgb}")
        except:
            pass # Ignore color errors

        # Construct tags
        result = text
        if run.font.bold:
            result = f"<b>{result}</b>"
        if run.font.italic:
            result = f"<i>{result}</i>"
        try:
            if run.font.underline:
                result = f"<u>{result}</u>"
        except: pass

        try:
            # Check for strikethrough (various properties)
            strike = False
            if hasattr(run.font, 'strike') and run.font.strike:
                strike = True
            if strike:
                result = f"<s>{result}</s>"
        except: pass

        attrs = []
        if style_parts:
            attrs.append(f'style="{"; ".join(style_parts)}"')

        try:
            if run.font.color.type == 2: # Theme
                attrs.append(f'data-pptx-theme-color="{run.font.color.theme_color}"')
        except: pass

        if attrs:
            result = f"<span {' '.join(attrs)}>{result}</span>"

        return result

    def _reconstruct_paragraph(self, paragraph, html_text):
        # Parse HTML
        parser = HTMLRunParser()
        try:
            parser.feed(html_text)
        except Exception as e:
            print(f"HTML Parse Error: {e} | Text: {html_text}")
            # Fallback: Just set text if parse fails?
            paragraph.clear()
            paragraph.add_run().text = html.unescape(html_text) # Strip tags implicitly? No.
            return

        parsed_runs = parser.runs
        if not parsed_runs:
            return

        # Clear and rebuild
        paragraph.clear()

        for p_run in parsed_runs:
            new_run = paragraph.add_run()
            new_run.text = html.unescape(p_run["text"])
            self._apply_style(new_run, p_run["style"])

    def _apply_style(self, run, style):
        run.font.bold = style["bold"]
        run.font.italic = style["italic"]
        run.font.underline = style["underline"]
        if style["strike"]:
            # Try setting strikethrough
            try:
                if hasattr(run.font, 'strike'):
                    run.font.strike = True
            except: pass

        if style["font_size"]:
            run.font.size = Pt(style["font_size"])

        if style["color_rgb"]:
            try:
                from pptx.dml.color import RGBColor
                run.font.color.rgb = RGBColor.from_string(style["color_rgb"])
            except: pass

        if style["theme_color"]:
            try:
                run.font.color.theme_color = int(style["theme_color"])
            except: pass

    def _calculate_max_chars(self, original_length):
        ratio = 1.0
        conf = self.translator.config

        s_lang = conf.source_language.lower()
        t_lang = conf.target_language.lower()
        base_ratio = conf.expansion_ratio

        if "japanese" in s_lang and "english" in t_lang:
            ratio = base_ratio
        elif "english" in s_lang and "japanese" in t_lang:
            ratio = 1.0 / base_ratio if base_ratio != 0 else 1.0

        return int(original_length * ratio)

    def save(self, output_path):
        self.prs.save(output_path)
