import os
import json
from openai import AzureOpenAI
from src.config import Config

class Translator:
    def __init__(self, config: Config, glossary: dict = None):
        self.config = config
        self.glossary = glossary or {}
        azure_conf = self.config.azure_openai

        # Initialize Azure OpenAI Client
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

    def translate_text(self, text: str, max_chars: int = None) -> str:
        """
        Translates the given text using Azure OpenAI.
        The text is expected to be formatted with XML-like tags.
        max_chars: Optional integer limit for the translated text length (excluding tags).
        """
        if not text or text.strip() == "":
            return text

        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": text}
        ]

        if max_chars is not None:
            messages.append({"role": "system", "content": f"IMPORTANT: Keep the total character count of the translated content (excluding tags) under {max_chars} characters. Do not remove or alter the tags <rN>."})

        try:
            response = self.client.chat.completions.create(
                model=self.deployment_name,
                messages=messages,
                temperature=0
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"Error during translation: {e}")
            return text

    def translate_batch(self, items: list) -> list:
        """
        Translates a batch of items.
        Input: list of {"id": int, "text": str, "max_chars": int}
        Output: list of {"id": int, "text": str}
        """
        if not items:
            return []

        # Construct JSON prompt
        prompt_items = []
        for item in items:
            prompt_items.append({
                "id": item["id"],
                "text": item["text"],
                "max_chars": item["max_chars"]
            })

        user_content = json.dumps(prompt_items, ensure_ascii=False)

        system_prompt = self.system_prompt + "\n\nProcess the following JSON list. Return a JSON list of objects with 'id' and 'text' (translated content). Maintain the same IDs."

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content}
        ]

        try:
            response = self.client.chat.completions.create(
                model=self.deployment_name,
                messages=messages,
                temperature=0,
                response_format={"type": "json_object"}
            )
            content = response.choices[0].message.content

            # The API might return a raw list or a wrapped object depending on how it interpreted "json_object".
            # Usually for "json_object" mode we need to instruct it to output JSON in the prompt (which we did).
            # But "json_object" enforces valid JSON.
            # We expect: {"translations": [...]} or just [...]
            # Let's try to parse
            try:
                parsed = json.loads(content)
                if isinstance(parsed, list):
                    return parsed
                elif isinstance(parsed, dict):
                    # Look for a list value
                    for val in parsed.values():
                        if isinstance(val, list):
                            return val
                    return [] # Fallback
                else:
                    return []
            except json.JSONDecodeError:
                print(f"Failed to parse JSON response: {content}")
                return []

        except Exception as e:
            error_str = str(e)
            print(f"Error during batch translation: {error_str}")

            # Handle Token Limit Exceeded
            # Azure OpenAI often returns "context_length_exceeded" in the error message or code.
            # We check for generic indication of length issues.
            if "context_length_exceeded" in error_str or "maximum context length" in error_str:
                if len(items) > 1:
                    print("Context length exceeded. Splitting batch...")
                    mid = len(items) // 2
                    left = items[:mid]
                    right = items[mid:]
                    return self.translate_batch(left) + self.translate_batch(right)
                else:
                    # Can't split further, just fail this item
                    print(f"Item too large to translate: {items[0]['id']}")
                    return []

            return []

class MockTranslator:
    def __init__(self, config: Config, glossary: dict = None):
        self.config = config

    def translate_text(self, text: str, max_chars: int = None) -> str:
        """
        Simulates translation by appending [EN] to content inside tags.
        Input: <r0>こんにちは</r0>
        Output: <r0>[EN] こんにちは</r0>
        """
        import re
        pattern = re.compile(r"(<r\d+>)(.*?)(</r\d+>)", re.DOTALL)
        def replace_match(match):
            tag_open, content, tag_close = match.groups()
            return f"{tag_open}[EN] {content}{tag_close}"
        return pattern.sub(replace_match, text)

    def translate_batch(self, items: list) -> list:
        """
        Simulates batch translation.
        """
        results = []
        for item in items:
            translated_text = self.translate_text(item["text"])
            results.append({
                "id": item["id"],
                "text": translated_text
            })
        return results
