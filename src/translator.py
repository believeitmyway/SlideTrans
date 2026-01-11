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
        self.base_system_prompt = self._build_base_system_prompt()

    def _build_base_system_prompt(self):
        # We don't construct the FULL prompt here anymore because we need to inject max_chars per request
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

        # Inject max_chars into the prompt
        if max_chars is not None:
            # Replace placeholder with actual number
            system_prompt = self.base_system_prompt.replace("{max_chars}", str(max_chars))
        else:
            # Remove the constraint sentence or replace with generic text
            # Assuming the prompt has "Keep the translation under {max_chars} characters..."
            # We can replace with "Keep the translation length reasonable." or just remove the specific number.
            # Simpler: replace with a very large number or just "unlimited" if the prompt flows well,
            # but better to replace the whole sentence?
            # Since the user specifically put the placeholder, we replace it.
            # If we just put "unlimited", the sentence "Keep the translation under unlimited characters" is weird.
            # Let's assume max_chars is usually provided. If not, we replace with "10000" or similar?
            # Or we can strip the sentence containing {max_chars}.
            # For now, let's replace with "appropriate" to keep it grammatical-ish if the prompt allows,
            # or simply "a reasonable number of".
            # Or better: "Keep the translation under 5000 characters" (a safe default).
            system_prompt = self.base_system_prompt.replace("{max_chars}", "5000")

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

        # Inject instruction for batching
        # We replace {max_chars} with a reference to the JSON field
        system_prompt = self.base_system_prompt.replace("{max_chars}", "the limit specified in the 'max_chars' field")

        system_prompt += "\n\nProcess the following JSON list. Return a JSON list of objects with 'id' and 'text' (translated content). Maintain the same IDs."

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

            try:
                parsed = json.loads(content)
                if isinstance(parsed, list):
                    return parsed
                elif isinstance(parsed, dict):
                    for val in parsed.values():
                        if isinstance(val, list):
                            return val
                    return []
                else:
                    return []
            except json.JSONDecodeError:
                print(f"Failed to parse JSON response: {content}")
                return []

        except Exception as e:
            error_str = str(e)
            print(f"Error during batch translation: {error_str}")

            if "context_length_exceeded" in error_str or "maximum context length" in error_str:
                if len(items) > 1:
                    print("Context length exceeded. Splitting batch...")
                    mid = len(items) // 2
                    left = items[:mid]
                    right = items[mid:]
                    return self.translate_batch(left) + self.translate_batch(right)
                else:
                    print(f"Item too large to translate: {items[0]['id']}")
                    return []

            return []

class MockTranslator:
    def __init__(self, config: Config, glossary: dict = None):
        self.config = config

    def translate_text(self, text: str, max_chars: int = None) -> str:
        """
        Simulates translation by appending [EN] to content inside tags.
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
