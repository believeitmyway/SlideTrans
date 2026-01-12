import re
from tqdm import tqdm
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

class PPTXProcessor:
    def __init__(self, filepath, translator):
        self.filepath = filepath
        self.translator = translator
        self.prs = Presentation(filepath)
        self.standard_tasks = []
        self.constrained_tasks = []

    def process(self):
        """
        Iterates through all slides and shapes, collecting text tasks,
        then processes them in batches.
        """
        self.standard_tasks = []
        self.constrained_tasks = []

        # Step 1: Collect all paragraphs that need translation
        for slide in self.prs.slides:
            for shape in slide.shapes:
                self._collect_tasks(shape, context="standard")

        # Step 2: Process collected tasks in batches
        # Process Standard Tasks
        if self.standard_tasks:
            self._process_batches(self.standard_tasks, "presentation_body_prompt", "Translating Standard Text")

        # Process Constrained Tasks (Tables/Groups)
        if self.constrained_tasks:
            self._process_batches(self.constrained_tasks, "constrained_text_prompt", "Translating Constrained Text")

    def _collect_tasks(self, shape, context="standard"):
        # Handle Groups (Recursive)
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for item in shape.shapes:
                # Items inside a group are constrained
                self._collect_tasks(item, context="constrained")
            return

        # Handle Tables
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if not cell.text_frame:
                        continue
                    # Items inside a table are constrained
                    self._collect_text_frame_tasks(cell.text_frame, context="constrained")
            return

        # Handle Text Frames
        if shape.has_text_frame:
            self._collect_text_frame_tasks(shape.text_frame, context=context)

    def _collect_text_frame_tasks(self, text_frame, context="standard"):
        for paragraph in text_frame.paragraphs:
            if not paragraph.text.strip():
                continue

            # Extract tagged text and calculate max chars immediately
            tagged_text, run_map = self._extract_tagged_text(paragraph)
            if not tagged_text:
                continue

            raw_text_length = len("".join([r.text for r in paragraph.runs]))
            max_chars = self._calculate_max_chars(raw_text_length)

            task = {
                "paragraph": paragraph,
                "tagged_text": tagged_text,
                "run_map": run_map,
                "max_chars": max_chars,
                "original_text": "".join([r.text for r in paragraph.runs])
            }

            if context == "constrained":
                self.constrained_tasks.append(task)
            else:
                self.standard_tasks.append(task)

    def _process_batches(self, tasks, prompt_config_key, desc):
        batch_size = self.translator.config.batch_size
        total_tasks = len(tasks)

        # Determine prompt template
        if prompt_config_key == "presentation_body_prompt":
            prompt_template = self.translator.config.presentation_body_prompt
        elif prompt_config_key == "constrained_text_prompt":
            prompt_template = self.translator.config.constrained_text_prompt
        else:
            prompt_template = None

        with tqdm(total=total_tasks, desc=desc, unit="task") as pbar:
            for i in range(0, total_tasks, batch_size):
                batch = tasks[i:i + batch_size]

                # Prepare batch for translator
                batch_input = []
                for idx, task in enumerate(batch):
                    batch_input.append({
                        "id": idx, # ID relative to the batch
                        "text": task["tagged_text"],
                        "max_chars": task["max_chars"]
                    })

                # Translate batch
                results = self.translator.translate_batch(batch_input, system_prompt_template=prompt_template)

                # Map results back to tasks
                if results:
                    results_map = {item.get("id"): item.get("text") for item in results}

                    for idx, task in enumerate(batch):
                        translated_text = results_map.get(idx)
                        if translated_text:
                            # Parse and reconstruct
                            parsed_segments = self._parse_tagged_text(translated_text)
                            self._reconstruct_paragraph(task["paragraph"], parsed_segments, task["run_map"])

                pbar.update(len(batch))

    def save(self, output_path):
        self.prs.save(output_path)

    # Legacy method _process_paragraph is no longer used directly but _reconstruct_paragraph is.
    # We kept _reconstruct_paragraph logic.

    def _extract_tagged_text(self, paragraph):
        """
        Converts paragraph runs to a tagged string like <r0>Hello</r0> <r1>World</r1>.
        Returns:
            tagged_text (str): The constructed string.
            run_map (dict): Map of ID (str) -> Original Run object.
        """
        tagged_parts = []
        run_map = {}

        current_id = 0
        for run in paragraph.runs:
            text = run.text
            run_id = str(current_id)
            run_map[run_id] = run
            tagged_parts.append(f"<r{run_id}>{text}</r{run_id}>")
            current_id += 1

        return "".join(tagged_parts), run_map

    def _parse_tagged_text(self, text):
        """
        Parses string like <r0>Hola</r0> <r1>Mundo</r1> into a list of (run_id, content).
        """
        pattern = re.compile(r"<r(\d+)>(.*?)</r\1>", re.DOTALL)
        matches = pattern.findall(text)
        return matches

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

        # Clear existing runs
        paragraph.clear()

        # Add new runs
        for run_id, content in parsed_segments:
            original_run = run_map.get(run_id)
            if not original_run:
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
            pass
