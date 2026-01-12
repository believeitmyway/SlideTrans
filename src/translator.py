import os
import json
import datetime
from openai import AzureOpenAI
from src.config import Config

class Translator:
    def __init__(self, config: Config, glossary: dict = None, debug_mode: bool = False):
        self.config = config
        self.glossary = glossary or {}
        self.debug_mode = debug_mode
        azure_conf = self.config.azure_openai

        # Initialize Azure OpenAI Client
        self.client = AzureOpenAI(
            api_key=azure_conf.get("api_key"),
            api_version=azure_conf.get("api_version"),
            azure_endpoint=azure_conf.get("endpoint")
        )
        self.deployment_name = azure_conf.get("deployment_name")

    def _log_debug(self, messages, response_content):
        if not self.debug_mode:
            return

        try:
            with open("llm_debug.log", "a", encoding="utf-8") as f:
                timestamp = datetime.datetime.now().isoformat()
                f.write(f"--- [{timestamp}] REQUEST ---\n")
                f.write(json.dumps(messages, ensure_ascii=False, indent=2))
                f.write("\n")
                f.write(f"--- [{timestamp}] RESPONSE ---\n")
                f.write(str(response_content))
                f.write("\n" + "="*80 + "\n")
        except Exception as e:
            print(f"Failed to write debug log: {e}")

    def translate_text(self, text: str, max_chars: int = None, system_prompt_template: str = None) -> str:
        """
        Translates the given text using Azure OpenAI.
        The text is expected to be formatted with HTML tags.
        max_chars: Optional integer limit for the translated text length (excluding tags).
        system_prompt_template: Optional prompt to use (overrides default).
        """
        if not text or text.strip() == "":
            return text

        # Prepare System Prompt
        base_prompt = system_prompt_template if system_prompt_template else self.config.translation_prompt

        # Inject max_chars
        limit_str = str(max_chars) if max_chars is not None else "reasonable limit"
        system_prompt = base_prompt.replace("{max_chars}", limit_str)

        # Inject Target Language (if placeholder exists)
        if "{target_language}" in system_prompt:
            system_prompt = system_prompt.replace("{target_language}", self.config.target_language)

        # Glossary
        if self.glossary:
            glossary_instruction = "\n\nUse the following glossary for translation:\n"
            for term, translation in self.glossary.items():
                glossary_instruction += f"- {term}: {translation}\n"
            system_prompt += glossary_instruction

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text}
        ]

        try:
            response = self.client.chat.completions.create(
                model=self.deployment_name,
                messages=messages,
                temperature=0
            )
            content = response.choices[0].message.content

            self._log_debug(messages, content)

            return content
        except Exception as e:
            print(f"Error during translation: {e}")
            return text

class MockTranslator:
    def __init__(self, config: Config, glossary: dict = None, debug_mode: bool = False):
        self.config = config
        self.debug_mode = debug_mode

    def _log_debug(self, messages, response_content):
        if not self.debug_mode:
            return

        try:
            with open("llm_debug.log", "a", encoding="utf-8") as f:
                timestamp = datetime.datetime.now().isoformat()
                f.write(f"--- [{timestamp}] REQUEST (MOCK) ---\n")
                f.write(json.dumps(messages, ensure_ascii=False, indent=2))
                f.write("\n")
                f.write(f"--- [{timestamp}] RESPONSE (MOCK) ---\n")
                f.write(str(response_content))
                f.write("\n" + "="*80 + "\n")
        except Exception as e:
            print(f"Failed to write debug log: {e}")

    def translate_text(self, text: str, max_chars: int = None, system_prompt_template: str = None) -> str:
        """
        Simulates translation by appending [EN] to content inside tags.
        """
        # Simple simulation: just append [EN] to text content (ignoring tags logic for regex simplicity)
        # Actually, let's try to preserve tags.

        # Since we use HTML now, let's just append [EN] to text outside tags?
        # A simple regex to find text outside of <...>
        import re
        # Find text between > and <
        # Or start and <
        # Or > and end

        # Simpler: Translate "Hello" -> "[EN] Hello"
        # <b>Hello</b> -> <b>[EN] Hello</b>

        def replace_text(match):
            content = match.group(2)
            if not content.strip():
                return match.group(0)
            return f"{match.group(1)}[EN] {content}{match.group(3)}"

        # Matches >text<
        result = re.sub(r"(>)([^<]+)(<)", replace_text, text)

        # Matches start...<
        result = re.sub(r"^([^<]+)(<)", lambda m: f"[EN] {m.group(1)}{m.group(2)}", result)

        # Matches >...end
        result = re.sub(r"(>)([^<]+)$", lambda m: f"{m.group(1)}[EN] {m.group(2)}", result)

        # If no tags, just wrap
        if "<" not in text:
            result = f"[EN] {text}"

        messages = [{"role": "user", "content": text}]
        self._log_debug(messages, result)

        return result
