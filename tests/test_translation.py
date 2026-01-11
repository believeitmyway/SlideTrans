import unittest
from unittest.mock import MagicMock, patch
from src.pptx_processor import PPTXProcessor
from src.translator import Translator
from src.config import Config
from pptx import Presentation
from pptx.util import Pt

class TestPPTXProcessor(unittest.TestCase):
    def setUp(self):
        # Create a mock config
        self.mock_config = MagicMock(spec=Config)
        self.mock_config.azure_openai = {}
        self.mock_config.translation_prompt = "Mock Prompt"
        self.mock_config.source_language = "Japanese"
        self.mock_config.target_language = "English"

        # Create a mock translator
        self.mock_translator = MagicMock(spec=Translator)

        # Patch Presentation to avoid loading file
        with patch("src.pptx_processor.Presentation") as MockPresentation:
            # Create a mock Presentation object structure
            # This is tricky because python-pptx objects are complex.
            # Instead of mocking the whole library, we will rely on checking the _process_paragraph logic directly
            # by passing a mock paragraph.
            self.processor = PPTXProcessor("dummy.pptx", self.mock_translator)
            # Mock the presentation to avoid file loading error
            self.processor.prs = MockPresentation.return_value

    def test_tag_preservation_and_formatting(self):
        # Setup Mock Paragraph and Runs
        mock_paragraph = MagicMock()

        # Run 0: "Hello " (Bold)
        run0 = MagicMock()
        run0.text = "Hello "
        run0.font.name = "Arial"
        run0.font.size = Pt(12)
        run0.font.bold = True
        run0.font.italic = False
        run0.font.color.type = None # Simplify color for this test

        # Run 1: "World" (Italic)
        run1 = MagicMock()
        run1.text = "World"
        run1.font.name = "Arial"
        run1.font.size = Pt(12)
        run1.font.bold = False
        run1.font.italic = True
        run1.font.color.type = None

        mock_paragraph.runs = [run0, run1]

        # Setup Translator Mock Response
        # Simulating JP translation: Hello -> こんにちは, World -> 世界
        # But keeping tags: <r0>こんにちは </r0><r1>世界</r1>
        self.mock_translator.translate_text.return_value = "<r0>こんにちは </r0><r1>世界</r1>"

        # We need to capture the runs added to the paragraph after it's cleared.
        # Since mock_paragraph.clear() will be called, and then add_run().
        # We need to mock add_run to return a new mock run that we can inspect.

        new_run0 = MagicMock()
        new_run0.font = MagicMock()
        new_run1 = MagicMock()
        new_run1.font = MagicMock()

        mock_paragraph.add_run.side_effect = [new_run0, new_run1]

        # Execute
        self.processor._process_paragraph(mock_paragraph)

        # Verify
        # 1. Verify Translator was called with correct tagged input
        # Input should be: <r0>Hello </r0><r1>World</r1>
        self.mock_translator.translate_text.assert_called_with("<r0>Hello </r0><r1>World</r1>")

        # 2. Verify Paragraph was cleared
        mock_paragraph.clear.assert_called_once()

        # 3. Verify New Runs have correct text and formatting
        self.assertEqual(new_run0.text, "こんにちは ")
        self.assertEqual(new_run0.font.bold, True) # Should inherit from run0
        self.assertEqual(new_run0.font.italic, False)

        self.assertEqual(new_run1.text, "世界")
        self.assertEqual(new_run1.font.bold, False)
        self.assertEqual(new_run1.font.italic, True) # Should inherit from run1

    def test_font_resizing(self):
        # Test case where translation is significantly longer
        mock_paragraph = MagicMock()

        # Original: "Hi" (Width = 2)
        run0 = MagicMock()
        run0.text = "Hi"
        run0.font.size = MagicMock() # Use MagicMock instead of actual Pt object to allow setting property
        run0.font.size.pt = 10.0

        mock_paragraph.runs = [run0]

        # Translated: "Hiiii" (Width = 5)
        # This should trigger resizing. Ratio = 2/5 = 0.4
        self.mock_translator.translate_text.return_value = "<r0>Hiiii</r0>"

        new_run0 = MagicMock()
        new_run0.font = MagicMock()
        # Mocking size assignment
        # The code does: target.font.size = Pt(source.font.size.pt * scaling_factor)

        mock_paragraph.add_run.return_value = new_run0

        # Execute
        self.processor._process_paragraph(mock_paragraph)

        # Verify
        # Check if Pt was called with reduced size
        # We can't easily check the exact Pt object equality, but we can check the call logic or inspect the assignment if we could.
        # Since Pt is a value object, let's assume if the logic runs, it's correct.
        # But we can verify that the code *calculated* the scaling factor correctly.
        # To do this, we can patch pptx.util.Pt to see what it was called with.

    @patch("src.pptx_processor.Pt")
    def test_font_resizing_logic(self, mock_pt):
        mock_paragraph = MagicMock()
        run0 = MagicMock()
        run0.text = "Hi"
        run0.font.size = MagicMock() # Mock the size object
        run0.font.size.pt = 10.0
        mock_paragraph.runs = [run0]

        self.mock_translator.translate_text.return_value = "<r0>Hiiii</r0>"

        new_run0 = MagicMock()
        new_run0.font = MagicMock()
        mock_paragraph.add_run.return_value = new_run0

        self.processor._process_paragraph(mock_paragraph)

        # Original Width ("Hi") = 2
        # Translated Width ("Hiiii") = 5
        # Expected Size = 10.0 * (2/5) = 4.0

        mock_pt.assert_called_with(4.0)

class TestTranslator(unittest.TestCase):
    @patch("src.translator.AzureOpenAI")
    def test_prompt_construction(self, mock_azure):
        config = MagicMock(spec=Config)
        config.azure_openai = {"api_key": "dummy", "endpoint": "dummy", "api_version": "dummy"}
        config.translation_prompt = "Base Prompt."
        config.source_language = "Japanese"
        config.target_language = "English"

        glossary = {"TermA": "TransA", "TermB": "TransB"}

        translator = Translator(config, glossary)

        expected_part_1 = "Base Prompt."
        expected_part_2 = "Translate from Japanese to English."
        expected_part_3 = "TermA: TransA"

        self.assertIn(expected_part_1, translator.system_prompt)
        self.assertIn(expected_part_2, translator.system_prompt)
        self.assertIn(expected_part_3, translator.system_prompt)
