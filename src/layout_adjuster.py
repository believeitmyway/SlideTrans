from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
from tqdm import tqdm

class LayoutAdjuster:
    def __init__(self, filepath):
        self.filepath = filepath
        self.prs = Presentation(filepath)

    def adjust(self):
        """
        Iterates through slides and adjusts text boxes and tables.
        """
        for slide_idx, slide in enumerate(tqdm(self.prs.slides, desc="Adjusting Layout")):
            self._adjust_slide(slide)

    def save(self, output_path):
        self.prs.save(output_path)

    def _adjust_slide(self, slide):
        # We need to collect all shapes first to do collision detection
        shapes = list(slide.shapes)

        for shape in shapes:
            if shape.has_text_frame:
                self._adjust_text_box(shape, shapes, slide.slide_layout)

            if shape.has_table:
                self._adjust_table(shape)

            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                self._adjust_group(shape)

    def _adjust_text_box(self, shape, all_shapes, layout):
        # Only adjust if it's not a table or chart
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            return

        try:
            current_left = shape.left
            current_top = shape.top
            current_width = shape.width
            current_height = shape.height
            current_right = current_left + current_width

            # Find nearest obstruction to the right
            slide_width = self.prs.slide_width
            max_right = slide_width

            # Check other shapes
            for other in all_shapes:
                if other == shape:
                    continue

                # Simple AABB collision check for "in the same horizontal band"
                if other.left >= current_right:
                    if not (other.top + other.height < current_top or other.top > current_top + current_height):
                        if other.left < max_right:
                            max_right = other.left

            # Calculate new width
            # Leave a small margin (e.g., 0.1 inch)
            margin = 91440
            available_width = max_right - current_left - margin

            # 1. Widen the shape if possible
            if available_width > current_width:
                shape.width = int(available_width)

            # Ensure word wrap is on (important for multi-line calculation)
            shape.text_frame.word_wrap = True

            # 2. Manual Font Scaling
            self._apply_manual_fit(shape.text_frame, shape.width)

        except Exception as e:
            # print(f"Error adjusting shape: {e}")
            pass

    def _adjust_table(self, shape):
        for row in shape.table.rows:
            for cell in row.cells:
                if cell.text_frame:
                    cell.text_frame.word_wrap = True
                    # Check against cell width
                    self._apply_manual_fit(cell.text_frame, cell.width)

    def _adjust_group(self, group_shape):
        for shape in group_shape.shapes:
             if shape.has_text_frame:
                 shape.text_frame.word_wrap = True
                 self._apply_manual_fit(shape.text_frame, shape.width)

    def _apply_manual_fit(self, text_frame, available_width):
        """
        Estimates text width and shrinks font if it exceeds available_width.
        """
        if not available_width or available_width <= 0:
            return

        text_width = self._estimate_text_width(text_frame)

        if text_width > available_width:
            ratio = available_width / text_width
            # Apply a small buffer to be safe (e.g., 95%)
            safe_ratio = ratio * 0.95
            self._scale_font_size(text_frame, safe_ratio)

    def _estimate_text_width(self, text_frame):
        """
        Estimates the visual width of the longest line in the text frame in EMUs.
        """
        max_line_width = 0
        current_line_width = 0

        # We need to handle paragraphs and possible soft breaks.
        # Ideally, we iterate runs.
        # But separate paragraphs implies new lines.

        for paragraph in text_frame.paragraphs:
            # Reset line width for new paragraph
            # But wait, what if the previous paragraph didn't end with a newline?
            # In PPT, paragraphs are block elements. They always start on a new line.
            if current_line_width > max_line_width:
                max_line_width = current_line_width
            current_line_width = 0

            for run in paragraph.runs:
                font_size = run.font.size
                if font_size is None:
                    font_size = Pt(18) # Default fallback

                # Calculate width of this run
                # Check for manual newlines in text (though normally run text doesn't contain \n unless explicitly added)
                text = run.text

                # Split by newline if present (preservation logic might put \n in runs)
                parts = text.split('\n')

                for i, part in enumerate(parts):
                    if i > 0:
                        # New line started
                        if current_line_width > max_line_width:
                            max_line_width = current_line_width
                        current_line_width = 0

                    # Calculate width of part
                    part_width = 0
                    for char in part:
                        # Estimate width: Wide ~ 1.0 * Size, Narrow ~ 0.5 * Size
                        # 1 pt = 12700 EMUs
                        is_wide = ord(char) > 255
                        factor = 1.0 if is_wide else 0.55 # 0.55 is a bit safer for variable width fonts
                        char_width = font_size.pt * 12700 * factor
                        part_width += char_width

                    current_line_width += part_width

        if current_line_width > max_line_width:
            max_line_width = current_line_width

        return max_line_width

    def _scale_font_size(self, text_frame, ratio):
        """
        Multiplies the font size of all runs by the ratio.
        """
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                if run.font.size:
                    new_size = run.font.size.pt * ratio
                    # Set minimum size limit? (e.g. 6pt)
                    if new_size < 6:
                        new_size = 6
                    run.font.size = Pt(new_size)
                else:
                    # If size was None (inherited), set it to scaled default
                    # Default 18pt * ratio
                    new_size = 18 * ratio
                    if new_size < 6:
                        new_size = 6
                    run.font.size = Pt(new_size)
