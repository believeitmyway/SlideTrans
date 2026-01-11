import argparse
import os
import sys
from src.config import Config
from src.translator import Translator
from src.pptx_processor import PPTXProcessor

def main():
    parser = argparse.ArgumentParser(description="Translate PowerPoint files using Azure OpenAI.")
    parser.add_argument("input_file", help="Path to the input .pptx file")
    parser.add_argument("--config", default="config.yaml", help="Path to configuration file")

    args = parser.parse_args()

    # Validation
    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' not found.")
        sys.exit(1)

    try:
        # Initialization
        print("Loading configuration...")
        config = Config(args.config)

        print("Initializing translator...")
        translator = Translator(config)

        print(f"Processing '{args.input_file}'...")
        processor = PPTXProcessor(args.input_file, translator)
        processor.process()

        # Save output
        filename, ext = os.path.splitext(args.input_file)
        output_file = f"{filename}_translated{ext}"
        processor.save(output_file)

        print(f"Done! Saved translated file to '{output_file}'.")

    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
