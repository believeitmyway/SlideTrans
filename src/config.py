import yaml
import os

class Config:
    def __init__(self, config_path="config.yaml"):
        self.config_path = config_path
        self._config = self._load_config()

    def _load_config(self):
        if not os.path.exists(self.config_path):
            raise FileNotFoundError(f"Config file not found: {self.config_path}")
        with open(self.config_path, "r", encoding="utf-8") as f:
            return yaml.safe_load(f)

    @property
    def azure_openai(self):
        return self._config.get("azure_openai", {})

    @property
    def translation_prompt(self):
        # Fallback for backward compatibility or default
        return self.presentation_body_prompt

    @property
    def presentation_body_prompt(self):
        return self._config.get("translation", {}).get("presentation_body_prompt", "")

    @property
    def constrained_text_prompt(self):
        return self._config.get("translation", {}).get("constrained_text_prompt", "")

    @property
    def source_language(self):
        return self._config.get("translation", {}).get("source_language", "Japanese")

    @property
    def target_language(self):
        return self._config.get("translation", {}).get("target_language", "English")

    @property
    def glossary_path(self):
        return self._config.get("translation", {}).get("glossary_path", "glossary.json")

    @property
    def expansion_ratio(self):
        return self._config.get("translation", {}).get("expansion_ratio", 1.0)

    @property
    def max_parallel_requests(self):
        return self._config.get("translation", {}).get("max_parallel_requests", 5)
