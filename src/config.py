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
        return self._config.get("translation", {}).get("prompt", "")
