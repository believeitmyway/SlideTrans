import os
from openai import AzureOpenAI
from src.config import Config

class Translator:
    def __init__(self, config: Config, glossary: dict = None):
        self.config = config
        self.glossary = glossary or {}
        azure_conf = self.config.azure_openai

        # Initialize Azure OpenAI Client
        # Note: In a real scenario, we might want to check if keys are present.
        # For testing with mocks, we just need the class structure.
        self.client = AzureOpenAI(
            api_key=azure_conf.get("api_key"),
            api_version=azure_conf.get("api_version"),
            azure_endpoint=azure_conf.get("endpoint")
        )
        self.deployment_name = azure_conf.get("deployment_name")
        self.system_prompt = self._build_system_prompt()

    def _build_system_prompt(self):
        base_prompt = self.config.translation_prompt
        lang_instruction = f" Translate from {self.config.source_language} to {self.config.target_language}."

        glossary_instruction = ""
        if self.glossary:
            glossary_instruction = "\n\nUse the following glossary for translation:\n"
            for term, translation in self.glossary.items():
                glossary_instruction += f"- {term}: {translation}\n"

        return base_prompt + lang_instruction + glossary_instruction

    def translate_text(self, text: str) -> str:
        """
        Translates the given text using Azure OpenAI.
        The text is expected to be formatted with XML-like tags.
        """
        if not text or text.strip() == "":
            return text

        try:
            response = self.client.chat.completions.create(
                model=self.deployment_name,
                messages=[
                    {"role": "system", "content": self.system_prompt},
                    {"role": "user", "content": text}
                ],
                temperature=0
            )
            return response.choices[0].message.content
        except Exception as e:
            # In a real app, we might log this.
            # For now, print to stderr or just re-raise.
            print(f"Error during translation: {e}")
            # If translation fails, return original text to avoid data loss
            return text

class MockTranslator:
    def __init__(self, config: Config, glossary: dict = None):
        self.config = config

    def translate_text(self, text: str) -> str:
        """
        Simulates translation by appending [EN] to content inside tags.
        Input: <r0>こんにちは</r0>
        Output: <r0>[EN] こんにちは</r0>
        """
        import re

        # Regex to find <rN>content</rN>
        pattern = re.compile(r"(<r\d+>)(.*?)(</r\d+>)", re.DOTALL)

        def replace_match(match):
            tag_open, content, tag_close = match.groups()
            return f"{tag_open}[EN] {content}{tag_close}"

        return pattern.sub(replace_match, text)
