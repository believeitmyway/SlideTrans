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
        # Kept for compatibility/fallback.
        if not text or text.strip() == "":
            return text

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
        Translates a batch of text items using plain text format:
        ID ::: LIMIT ::: TEXT
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

        # Prepare User Content (Text Format)
        # item keys: id, text, limit
        lines = []
        for item in items:
            line = f"{item['id']} ::: {item['limit']} ::: {item['text']}"
            lines.append(line)
        user_content = "\n".join(lines)

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

            # Parse Text Response
            # Expected format: ID ::: TRANSLATION
            translated_items = []
            for line in content.splitlines():
                if ":::" not in line:
                    continue
                parts = line.split(":::", 1)
                if len(parts) < 2:
                    continue

                try:
                    t_id = int(parts[0].strip())
                    t_text = parts[1].strip()
                    translated_items.append({"id": t_id, "translation": t_text})
                except ValueError:
                    continue

            return translated_items

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
        # Regex-based simple translation simulation
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
        Simulates batch translation with text format.
        """
        translated_items = []
        response_lines = []

        for item in items:
            t_text = self.translate_text(item["text"])
            translated_items.append({
                "id": item["id"],
                "translation": t_text
            })
            response_lines.append(f"{item['id']} ::: {t_text}")

        # Log effectively what we would get back
        self._log_debug(
            [{"role": "user", "content": "MOCKED BATCH REQUEST"}],
            "\n".join(response_lines)
        )
        return translated_items
