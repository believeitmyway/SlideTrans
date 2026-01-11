import re
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

class PPTXProcessor:
    def __init__(self, filepath, translator):
        self.filepath = filepath
        self.translator = translator
        self.prs = Presentation(filepath)

    def process(self):
        """
        Iterates through all slides and shapes, translating text.
        """
        for slide in self.prs.slides:
            for shape in slide.shapes:
                self._process_shape(shape)

    def _process_shape(self, shape):
        # Handle Groups (Recursive)
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for item in shape.shapes:
                self._process_shape(item)
            return

        # Handle Tables
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if not cell.text_frame:
                        continue
                    self._process_text_frame(cell.text_frame)
            return

        # Handle Text Frames
        if shape.has_text_frame:
            self._process_text_frame(shape.text_frame)

    def _process_text_frame(self, text_frame):
        for paragraph in text_frame.paragraphs:
            if not paragraph.text.strip():
                continue
            self._process_paragraph(paragraph)

    def save(self, output_path):
        self.prs.save(output_path)

    def _process_paragraph(self, paragraph):
        # 1. Extract runs and build tagged text
        tagged_text, run_map = self._extract_tagged_text(paragraph)

        if not tagged_text:
            return

        # Calculate char limit
        # Pure text length
        raw_text_length = len("".join([r.text for r in paragraph.runs]))
        max_chars = self._calculate_max_chars(raw_text_length)
        # Note: We pass max_chars to translate_text. The prompt construction logic needs to be updated.
        # But wait, translate_text signature needs update.

        # 2. Translate
        translated_text = self.translator.translate_text(tagged_text, max_chars=max_chars)

        # 3. Parse translated text
        parsed_segments = self._parse_tagged_text(translated_text)

        # 4. Reconstruct paragraph
        self._reconstruct_paragraph(paragraph, parsed_segments, run_map)

    def _extract_tagged_text(self, paragraph):
        """
        Converts paragraph runs to a tagged string like <r0>Hello</r0> <r1>World</r1>.
        Returns:
            tagged_text (str): The constructed string.
            run_map (dict): Map of ID (str) -> Original Run object.
        """
        tagged_parts = []
        run_map = {}

        # We filter out empty runs to avoid cluttering the prompt,
        # unless they contain meaningful whitespace, but usually empty runs are artifacts.
        current_id = 0
        for run in paragraph.runs:
            # We preserve all runs to maintain spacing if they have text
            text = run.text
            # Note: We don't strip text here because spaces are important.

            run_id = str(current_id)
            run_map[run_id] = run

            # Escape XML special characters in text to avoid confusing the tag parser later?
            # Ideally yes, but for now assuming simple text or that LLM handles it.
            # Let's do basic escaping if needed, but < > might confuse regex.
            # We'll use a unique tag delimiter that is unlikely to be in text if we were rigorous,
            # but <rN> is requested.

            tagged_parts.append(f"<r{run_id}>{text}</r{run_id}>")
            current_id += 1

        return "".join(tagged_parts), run_map

    def _parse_tagged_text(self, text):
        """
        Parses string like <r0>Hola</r0> <r1>Mundo</r1> into a list of (run_id, content).
        """
        # Regex to find <rN>content</rN>
        # Non-greedy match for content
        pattern = re.compile(r"<r(\d+)>(.*?)</r\1>", re.DOTALL)
        matches = pattern.findall(text)
        return matches

    def _calculate_max_chars(self, original_length):
        # Logic:
        # JP -> EN (Source=JP, Target=EN): Ratio (e.g. 1.7)
        # EN -> JP (Source=EN, Target=JP): 1/Ratio (e.g. 0.58)
        # Other: 1.0

        ratio = 1.0
        conf = self.translator.config # Access config via translator

        s_lang = conf.source_language.lower()
        t_lang = conf.target_language.lower()

        base_ratio = conf.expansion_ratio

        if "japanese" in s_lang and "english" in t_lang:
            ratio = base_ratio
        elif "english" in s_lang and "japanese" in t_lang:
            ratio = 1.0 / base_ratio if base_ratio != 0 else 1.0

        return int(original_length * ratio)

    def _reconstruct_paragraph(self, paragraph, parsed_segments, run_map):
        if not parsed_segments:
            return

        # Calculate width scaling factor
        original_text = "".join([r.text for r in paragraph.runs])
        translated_text = "".join([content for _, content in parsed_segments])

        orig_width = self._estimate_width(original_text)
        trans_width = self._estimate_width(translated_text)

        scaling_factor = 1.0
        if trans_width > orig_width and orig_width > 0:
            scaling_factor = orig_width / trans_width
            # Limit scaling to not be too tiny? User said "automatically reduce",
            # let's assume no lower bound for now, or maybe 0.5 minimum.
            # But the user wants layout preserved, so fitting is priority.

        # Clear existing runs
        # Note: p.clear() removes all runs.
        paragraph.clear()

        # Add new runs
        for run_id, content in parsed_segments:
            original_run = run_map.get(run_id)
            if not original_run:
                # If AI hallucinated a new ID, we skip or use default style.
                # Let's create a run with default style (inherit from paragraph)
                new_run = paragraph.add_run()
            else:
                new_run = paragraph.add_run()
                self._copy_run_formatting(original_run, new_run, scaling_factor)

            new_run.text = content

    def _estimate_width(self, text):
        """
        Estimates visual width.
        Wide characters (East Asian) = 2
        Narrow characters (ASCII) = 1
        """
        width = 0
        for char in text:
            # Simple check: if ord(char) > 255, assume wide.
            # This is a rough heuristic.
            if ord(char) > 255:
                width += 2
            else:
                width += 1
        return width

    def _copy_run_formatting(self, source, target, scaling_factor=1.0):
        # Font properties
        if source.font.name:
            target.font.name = source.font.name

        # Size
        if source.font.size:
            target.font.size = Pt(source.font.size.pt * scaling_factor)

        # Bold/Italic
        target.font.bold = source.font.bold
        target.font.italic = source.font.italic

        # Color
        try:
            if source.font.color.type:
                if source.font.color.type == 1: # RGB
                    target.font.color.rgb = source.font.color.rgb
                elif source.font.color.type == 2: # Theme
                    target.font.color.theme_color = source.font.color.theme_color
                    target.font.color.brightness = source.font.color.brightness
        except Exception:
            # Color copying can be tricky if properties are not set
            pass
