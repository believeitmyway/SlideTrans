import argparse
import json
import os
import sys
from src.config import Config
from src.translator import Translator, MockTranslator
from src.pptx_processor import PPTXProcessor
from src.layout_adjuster import LayoutAdjuster

def load_glossary(glossary_path):
    if not os.path.exists(glossary_path):
        print(f"Glossary file '{glossary_path}' not found. Skipping.")
        return {}
    try:
        with open(glossary_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading glossary: {e}. Skipping.")
        return {}

def main():
    parser = argparse.ArgumentParser(description="Translate PowerPoint files using Azure OpenAI.")
    parser.add_argument("input_file", help="Path to the input .pptx file")
    parser.add_argument("--config", default="config.yaml", help="Path to configuration file")
    parser.add_argument("--mock", action="store_true", help="Use mock translator without API calls")
    parser.add_argument("--debug-llm", action="store_true", help="Log LLM prompts and responses to a file")
    parser.add_argument("--output", help="Path to the output .pptx file")

    args = parser.parse_args()

    # Validation
    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' not found.")
        sys.exit(1)

    try:
        # Initialization
        print("Loading configuration...")
        config = Config(args.config)

        # Load glossary
        glossary = load_glossary(config.glossary_path)

        print("Initializing translator...")
        if args.mock:
            print("Using Mock Translator.")
            translator = MockTranslator(config, glossary, debug_mode=args.debug_llm)
        else:
            translator = Translator(config, glossary, debug_mode=args.debug_llm)

        print(f"Processing '{args.input_file}'...")
        processor = PPTXProcessor(args.input_file, translator)
        processor.process()

        # Intermediate save (optional, but good for safety)
        filename, ext = os.path.splitext(args.input_file)
        intermediate_file = f"{filename}_translated_raw{ext}"
        processor.save(intermediate_file)

        print("Adjusting layout...")
        # Layout Adjustment Phase
        adjuster = LayoutAdjuster(intermediate_file)
        adjuster.adjust()

        # Determine Output File
        if args.output:
            output_file = args.output
        else:
            output_file = f"{filename}_translated{ext}"

        adjuster.save(output_file)

        # Clean up intermediate file
        if os.path.exists(intermediate_file):
            os.remove(intermediate_file)

        print(f"Done! Saved translated file to '{output_file}'.")

    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
