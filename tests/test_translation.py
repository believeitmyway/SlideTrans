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
        self.mock_translator.config = self.mock_config # Attach config to translator mock

        # Patch Presentation to avoid loading file
        with patch("src.pptx_processor.Presentation") as MockPresentation:
            self.processor = PPTXProcessor("dummy.pptx", self.mock_translator)
            self.processor.prs = MockPresentation.return_value

    def test_tag_preservation_and_formatting(self):
        # Setup Mock Paragraph and Runs
        mock_paragraph = MagicMock()

        # Run 0: "Hello " (Bold)
        run0 = MagicMock(); run0.text = "Hello "; run0.font.bold = True; run0.font.italic = False; run0.font.color.type = None
        # Run 1: "World" (Italic)
        run1 = MagicMock(); run1.text = "World"; run1.font.bold = False; run1.font.italic = True; run1.font.color.type = None

        mock_paragraph.runs = [run0, run1]

        # Manually create run_map as it would be created by extract_tagged_text
        run_map = {"0": run0, "1": run1}

        # Simulate translated parsed segments
        # <r0>こんにちは </r0><r1>世界</r1>
        parsed_segments = [("0", "こんにちは "), ("1", "世界")]

        # Mock add_run
        new_run0 = MagicMock(); new_run0.font = MagicMock()
        new_run1 = MagicMock(); new_run1.font = MagicMock()
        mock_paragraph.add_run.side_effect = [new_run0, new_run1]

        # Execute
        self.processor._reconstruct_paragraph(mock_paragraph, parsed_segments, run_map)

        # Verify
        mock_paragraph.clear.assert_called_once()

        # Verify New Runs
        self.assertEqual(new_run0.text, "こんにちは ")
        self.assertEqual(new_run0.font.bold, True)

        self.assertEqual(new_run1.text, "世界")
        self.assertEqual(new_run1.font.italic, True)

    def test_font_resizing(self):
        # Setup Mock Paragraph
        mock_paragraph = MagicMock()
        run0 = MagicMock(); run0.text = "Hi"; run0.font.size.pt = 10.0
        mock_paragraph.runs = [run0]
        run_map = {"0": run0}

        # Translated: "Hiiii" (Width 5) vs "Hi" (Width 2) -> Ratio 0.4
        parsed_segments = [("0", "Hiiii")]

        new_run0 = MagicMock(); new_run0.font = MagicMock()
        mock_paragraph.add_run.return_value = new_run0

        # Execute
        self.processor._reconstruct_paragraph(mock_paragraph, parsed_segments, run_map)

        # Verify logic implicitly by checking calls if we could, but here we can at least ensure it ran without error
        # To strictly verify resizing, we rely on the next test which mocks Pt
        mock_paragraph.clear.assert_called_once()

    @patch("src.pptx_processor.Pt")
    def test_font_resizing_logic(self, mock_pt):
        mock_paragraph = MagicMock()
        run0 = MagicMock(); run0.text = "Hi"; run0.font.size.pt = 10.0
        mock_paragraph.runs = [run0]
        run_map = {"0": run0}

        # Translated: "Hiiii" (Width 5) vs "Hi" (Width 2) -> Scaling 0.4
        parsed_segments = [("0", "Hiiii")]

        new_run0 = MagicMock(); new_run0.font = MagicMock()
        mock_paragraph.add_run.return_value = new_run0

        # We must pass context="constrained" to trigger font resizing
        self.processor._reconstruct_paragraph(mock_paragraph, parsed_segments, run_map, context="constrained")

        # Expected Size = 10.0 * 0.4 = 4.0
        mock_pt.assert_called_with(4.0)

    def test_calculate_max_chars(self):
        # Case 1: Japanese to English (Ratio 1.7)
        self.mock_config.source_language = "Japanese"
        self.mock_config.target_language = "English"
        self.mock_config.expansion_ratio = 1.7
        self.mock_translator.config = self.mock_config

        self.assertEqual(self.processor._calculate_max_chars(10), 17)

        # Case 2: English to Japanese (Ratio 1/1.7 = 0.588)
        self.mock_config.source_language = "English"
        self.mock_config.target_language = "Japanese"
        self.assertEqual(self.processor._calculate_max_chars(100), int(100 * (1/1.7)))

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

        self.assertIn("Base Prompt.", translator.base_system_prompt)
        self.assertIn("Translate from Japanese to English.", translator.base_system_prompt)
        self.assertIn("TermA: TransA", translator.base_system_prompt)
