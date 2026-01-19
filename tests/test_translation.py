import unittest
import json
from unittest.mock import MagicMock, patch
from src.pptx_processor import PPTXProcessor, HTMLRunParser
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
        self.mock_config.expansion_ratio = 1.7
        self.mock_config.max_parallel_requests = 1

        # Create a mock translator
        self.mock_translator = MagicMock(spec=Translator)
        self.mock_translator.config = self.mock_config

        # Patch Presentation
        with patch("src.pptx_processor.Presentation") as MockPresentation:
            self.processor = PPTXProcessor("dummy.pptx", self.mock_translator)
            self.processor.prs = MockPresentation.return_value

    def test_run_to_html_spaces(self):
        run = MagicMock()
        run.text = " Hello "
        run.font.bold = True
        run.font.italic = False
        run.font.underline = False
        run.font.strike = False
        run.font.size.pt = None
        run.font.color.type = None

        html_out = self.processor._run_to_html(run)
        self.assertIn("<sp/>", html_out)
        self.assertIn("<b>", html_out)

    def test_reconstruct_paragraph_sp_tag(self):
        mock_paragraph = MagicMock()
        xml_text = '<b><sp/>Word<sp/></b>'

        added_runs = []
        def add_run_side_effect():
            r = MagicMock()
            r.font = MagicMock()
            added_runs.append(r)
            return r
        mock_paragraph.add_run.side_effect = add_run_side_effect

        self.processor._reconstruct_paragraph(mock_paragraph, xml_text)

        self.assertEqual(len(added_runs), 3)
        self.assertEqual(added_runs[0].text, " ")
        self.assertTrue(added_runs[0].font.bold)

        self.assertEqual(added_runs[1].text, "Word")

        self.assertEqual(added_runs[2].text, " ")

class TestTranslator(unittest.TestCase):
    @patch("src.translator.AzureOpenAI")
    def test_translate_batch_xml(self, mock_azure):
        config = MagicMock(spec=Config)
        config.azure_openai = {"api_key": "dummy", "endpoint": "dummy", "api_version": "dummy"}
        config.target_language = "English"

        # Mock Config Properties
        config.presentation_body_prompt = "Prompt"

        translator = Translator(config)

        mock_response = MagicMock()
        mock_xml_output = """<list><item id="0">T1</item></list>"""
        mock_response.choices[0].message.content = mock_xml_output
        translator.client.chat.completions.create.return_value = mock_response

        items = [{"id": 0, "text": "S1", "limit": 10}]
        result = translator.translate_batch(items, system_prompt_template="Prompt")
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]['translation'], "T1")
