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
        self.mock_config.presentation_body_prompt = "Mock Body Prompt"
        self.mock_config.constrained_text_prompt = "Mock Constrained Prompt"

        # Create a mock translator
        self.mock_translator = MagicMock(spec=Translator)
        self.mock_translator.config = self.mock_config # Attach config to translator mock

        # Patch Presentation to avoid loading file
        with patch("src.pptx_processor.Presentation") as MockPresentation:
            self.processor = PPTXProcessor("dummy.pptx", self.mock_translator)
            self.processor.prs = MockPresentation.return_value

    def test_run_to_html(self):
        # Mock Run
        run = MagicMock()
        run.text = "Hello"
        run.font.bold = True
        run.font.italic = False
        run.font.underline = False
        run.font.strike = False  # Explicitly set strike to False
        run.font.size.pt = None
        run.font.color.type = None

        # NOTE: When using MagicMock for pptx objects, accessing properties often returns a new MagicMock unless specified.
        # Ensure 'strike' resolves to False-y, not a Mock object (which is truthy).
        # We need to ensure getattr(run.font, 'strike') returns False.
        # But run.font.strike = False is safer.

        html_out = self.processor._run_to_html(run)
        self.assertEqual(html_out, "<b>Hello</b>")

    def test_reconstruct_paragraph_html(self):
        # Setup Mock Paragraph
        mock_paragraph = MagicMock()

        # HTML Text: <span style="font-size:12pt"><b>BoldText</b></span>
        html_text = '<span style="font-size:12pt"><b>BoldText</b></span>'

        # New Run Mock
        new_run = MagicMock()
        new_run.font = MagicMock()
        mock_paragraph.add_run.return_value = new_run

        # Execute
        self.processor._reconstruct_paragraph(mock_paragraph, html_text)

        # Verify
        mock_paragraph.clear.assert_called_once()
        self.assertEqual(new_run.text, "BoldText")
        self.assertEqual(new_run.font.bold, True)

        # Verify size setting (we need to mock Pt or check call args)
        # In implementation: run.font.size = Pt(style["font_size"])
        # Since we didn't patch Pt here, it will try to call real Pt.
        # But we haven't imported Pt in the test file scope effectively for the module being tested unless we patch it inside module.
        # Actually pptx.util.Pt is imported in src.pptx_processor.
        # We can check if new_run.font.size was set.

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
    def test_translate_text(self, mock_azure):
        config = MagicMock(spec=Config)
        config.azure_openai = {"api_key": "dummy", "endpoint": "dummy", "api_version": "dummy"}
        config.translation_prompt = "Translate this: {max_chars}"
        config.source_language = "Japanese"
        config.target_language = "English"

        glossary = {"TermA": "TransA"}

        translator = Translator(config, glossary)

        # Mock API Response
        mock_response = MagicMock()
        mock_response.choices[0].message.content = "Translated Text"
        translator.client.chat.completions.create.return_value = mock_response

        # Execute
        result = translator.translate_text("Source Text", max_chars=100)

        self.assertEqual(result, "Translated Text")

        # Verify call arguments
        call_args = translator.client.chat.completions.create.call_args
        messages = call_args.kwargs['messages']
        system_prompt = messages[0]['content']

        self.assertIn("100", system_prompt) # max_chars injected
        self.assertIn("TermA: TransA", system_prompt) # Glossary injected
