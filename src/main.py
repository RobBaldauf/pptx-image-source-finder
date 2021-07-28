import argparse
import os
import sys
from typing import Tuple

from image_resolver import ImageResolver
from result_writer import ResultWriter


def parse_args(args) -> Tuple[str, str, str]:
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", help="Powerpoint file to find image sources for", type=str, required=True)
    parser.add_argument("-o", "--output", help="Output Powerpoint file", type=str, required=True)
    parser.add_argument("-k", "--key-file", help="Google cloud vision key file", type=str, required=True)
    args = parser.parse_args(args)

    input_filename = os.path.abspath(args.input)
    output_filename = os.path.abspath(args.output)
    key_filename = os.path.abspath(args.key_file)
    if not os.path.isfile(input_filename):
        raise ValueError(f"Input file {input_filename} does not exist...")
    if not os.path.isfile(key_filename):
        raise ValueError(f"Key file {key_filename} does not exist...")
    return input_filename, output_filename, key_filename


def main(args) -> int:
    input_filename, output_filename, key_filename = parse_args(args)
    print(f"Looking up sources for images in powerpoint '{input_filename}'")
    resolver = ImageResolver(key_filename)
    slides = resolver.run(input_filename)
    print(f"Found {len(slides)} slides with images")
    if len(slides) == 0:
        print("Skipping output file creation...")
    else:
        writer = ResultWriter()
        print(f"Writing results to new powerpoint {output_filename}")
        writer.run(slides, output_filename)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
