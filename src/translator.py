import os
import json
import datetime
import re
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
        # Kept for compatibility if needed, though we primarily use translate_batch now.
        if not text or text.strip() == "":
            return text

        # This method assumes singular translation, might not be used in new flow
        # But leaving it intact just in case
        base_prompt = system_prompt_template if system_prompt_template else self.config.presentation_body_prompt
        limit_str = str(max_chars) if max_chars is not None else "reasonable limit"
        system_prompt = base_prompt.replace("{max_chars}", limit_str)

        if "{target_language}" in system_prompt:
            system_prompt = system_prompt.replace("{target_language}", self.config.target_language)

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

    def translate_batch(self, items: list, system_prompt_template: str) -> list:
        """
        Translates a batch of text items (JSON list).
        items: List of dicts [{"id":..., "text":..., "limit":...}]
        """
        if not items:
            return []

        # Prepare System Prompt
        system_prompt = system_prompt_template
        if "{target_language}" in system_prompt:
            system_prompt = system_prompt.replace("{target_language}", self.config.target_language)

        if self.glossary:
            glossary_instruction = "\n\nUse the following glossary for translation:\n"
            for term, translation in self.glossary.items():
                glossary_instruction += f"- {term}: {translation}\n"
            system_prompt += glossary_instruction

        # Prepare User Content (JSON)
        user_content = json.dumps(items, ensure_ascii=False)

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content}
        ]

        try:
            response = self.client.chat.completions.create(
                model=self.deployment_name,
                messages=messages,
                temperature=0
            )
            content = response.choices[0].message.content
            self._log_debug(messages, content)

            # Parse JSON
            try:
                # The LLM should return a JSON object containing the array, or just the array.
                # Since we asked for a JSON array in the prompt, it might be wrapped or just the list.
                # However, with response_format={"type": "json_object"}, the model is forced to generate a valid JSON object.
                # The prompt asks for a JSON array.
                # Note: 'json_object' mode requires the output to be a valid JSON object (dict), not list.
                # If the prompt asks for an array, 'json_object' mode might complain or force a wrapper.
                # Actually, standard JSON mode usually expects a root object {}.
                # Let's check the prompt again. I asked for a JSON array.
                # If I use response_format={"type": "json_object"}, I must ensure the prompt asks for a JSON object.
                # Or I can remove response_format constraint if I want a raw list.
                # Given the user wants "Simple", maybe I should just rely on text output and parse it.
                # But let's try to be robust.
                # Let's NOT use response_format={"type": "json_object"} if we want a list.
                # I will remove response_format to allow a top-level array.
                parsed = json.loads(content)
                if isinstance(parsed, dict) and "translations" in parsed:
                    # Handle case where LLM wraps it
                    return parsed["translations"]
                if isinstance(parsed, list):
                    return parsed
                # If it's a dict but we expected a list, maybe it wrapped it differently?
                print(f"Unexpected JSON structure: {type(parsed)}")
                return []
            except json.JSONDecodeError:
                print(f"Failed to decode JSON response: {content}")
                return []

        except Exception as e:
            print(f"Error during batch translation: {e}")
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

    def translate_text(self, text: str, max_chars: int = None, system_prompt_template: str = None) -> str:
        # Matches >text<
        def replace_text(match):
            content = match.group(2)
            if not content.strip():
                return match.group(0)
            return f"{match.group(1)}[EN] {content}{match.group(3)}"

        result = re.sub(r"(>)([^<]+)(<)", replace_text, text)
        result = re.sub(r"^([^<]+)(<)", lambda m: f"[EN] {m.group(1)}{m.group(2)}", result)
        result = re.sub(r"(>)([^<]+)$", lambda m: f"{m.group(1)}[EN] {m.group(2)}", result)
        if "<" not in text:
            result = f"[EN] {text}"
        return result

    def translate_batch(self, items: list, system_prompt_template: str) -> list:
        """
        Simulates batch translation.
        """
        translated_items = []
        for item in items:
            t_text = self.translate_text(item["text"])
            translated_items.append({
                "id": item["id"],
                "translation": t_text
            })

        self._log_debug([{"role": "user", "content": json.dumps(items)}], json.dumps(translated_items))
        return translated_items
