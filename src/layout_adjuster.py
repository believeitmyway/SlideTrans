from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE
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

        # Sort shapes by x position to help find neighbors?
        # A simple O(N^2) check is fine for typical slide shape counts (~10-50).

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

        # Strategy:
        # 1. Calculate current bounding box.
        # 2. Look for space to the right.
        # 3. Expand width.
        # 4. Set AutoFit.

        try:
            current_left = shape.left
            current_top = shape.top
            current_width = shape.width
            current_height = shape.height
            current_right = current_left + current_width

            # Find nearest obstruction to the right
            # Slide width constraint
            slide_width = self.prs.slide_width
            max_right = slide_width

            # Check other shapes
            for other in all_shapes:
                if other == shape:
                    continue

                # Simple AABB collision check for "in the same horizontal band"
                # If other is to the right of current
                if other.left >= current_right:
                    # And vertically overlaps
                    if not (other.top + other.height < current_top or other.top > current_top + current_height):
                        # It is a candidate obstruction
                        if other.left < max_right:
                            max_right = other.left

            # Calculate new width
            # Leave a small margin (e.g., 10px or 0.1 inch)
            margin = 91440 # 1 inch = 914400 EMUs. 0.1 inch margin.
            available_width = max_right - current_left - margin

            if available_width > current_width:
                shape.width = int(available_width)

            # Enable AutoFit "Shrink text on overflow"
            # TEXT_TO_FIT_SHAPE = 1

            # Ensure word wrap is on
            shape.text_frame.word_wrap = True

            # Clear explicit font sizes to force AutoFit recalculation
            self._clear_explicit_font_sizing(shape.text_frame)

            shape.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        except Exception as e:
            # print(f"Error adjusting shape: {e}")
            pass

    def _adjust_table(self, shape):
        # For tables, we do not change width.
        # We strictly want to ensure text fits.
        # python-pptx support for table auto-fit is limited.
        # We can try setting the property on the text frame of each cell.

        for row in shape.table.rows:
            for cell in row.cells:
                if cell.text_frame:
                    cell.text_frame.word_wrap = True
                    self._clear_explicit_font_sizing(cell.text_frame)
                    cell.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    def _adjust_group(self, group_shape):
        for shape in group_shape.shapes:
             if shape.has_text_frame:
                 # Recursive adjustment?
                 # Collision detection inside groups is hard because coordinates might be relative or transformed.
                 # Safest bet: Just set AutoFit.
                 shape.text_frame.word_wrap = True
                 self._clear_explicit_font_sizing(shape.text_frame)
                 shape.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    def _clear_explicit_font_sizing(self, text_frame):
        """
        Clears explicit font sizes from all runs in the text frame.
        This allows PowerPoint's AutoFit engine to take over control of the font size.
        """
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = None
