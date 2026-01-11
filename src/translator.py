import os
from openai import AzureOpenAI
from src.config import Config

class Translator:
    def __init__(self, config: Config):
        self.config = config
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
        self.prompt = self.config.translation_prompt

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
                    {"role": "system", "content": self.prompt},
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
