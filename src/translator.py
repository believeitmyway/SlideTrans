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
            content = response.choices[0].message.content

            self._log_debug(messages, content)

            return content
        except Exception as e:
            print(f"Error during translation: {e}")
            return text

    def translate_batch(self, items: list, system_prompt_template: str = None) -> list:
        """
        Translates a batch of items.
        Input: list of {"id": int, "text": str, "max_chars": int}
        Output: list of {"id": int, "text": str}
        system_prompt_template: Optional string to override the default prompt.
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

        # Determine base prompt to use
        base_prompt_to_use = system_prompt_template if system_prompt_template else self.base_system_prompt

        # Helper logic inline:
        if system_prompt_template:
            # We need to append the standard suffix (lang + glossary)
            # This duplicates logic from _build_base_system_prompt.
            # Let's extract the suffix.
            lang_instruction = f" Translate from {self.config.source_language} to {self.config.target_language}."
            glossary_instruction = ""
            if self.glossary:
                glossary_instruction = "\n\nUse the following glossary for translation:\n"
                for term, translation in self.glossary.items():
                    glossary_instruction += f"- {term}: {translation}\n"

            current_prompt = system_prompt_template + lang_instruction + glossary_instruction
        else:
            current_prompt = self.base_system_prompt

        # Inject instruction for batching
        system_prompt = current_prompt.replace("{max_chars}", "the limit specified in the 'max_chars' field")

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

            self._log_debug(messages, content)

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

    def translate_text(self, text: str, max_chars: int = None) -> str:
        """
        Simulates translation by appending [EN] to content inside tags.
        """
        import re
        pattern = re.compile(r"(<r\d+>)(.*?)(</r\d+>)", re.DOTALL)
        def replace_match(match):
            tag_open, content, tag_close = match.groups()
            return f"{tag_open}[EN] {content}{tag_close}"

        result = pattern.sub(replace_match, text)

        # Log simulated activity
        messages = [{"role": "user", "content": text}]
        self._log_debug(messages, result)

        return result

    def translate_batch(self, items: list, system_prompt_template: str = None) -> list:
        """
        Simulates batch translation.
        """
        results = []
        user_content_sim = json.dumps(items) # Simplified log payload

        for item in items:
            translated_text = self.translate_text(item["text"])
            results.append({
                "id": item["id"],
                "text": translated_text
            })

        # translate_text logs individually, but batch logic usually logs the batch request.
        # To avoid double logging or missing the batch structure, let's log the batch level here
        # and suppress individual logs? Or just log both.
        # Since Mock is just for testing, let's just log the batch structure to verify the format.

        messages = [{"role": "system", "content": "Mock Batch"}, {"role": "user", "content": user_content_sim}]
        self._log_debug(messages, json.dumps(results))

        return results
