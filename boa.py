import argparse
import sys
from oabparser import OABParser

def main():
  parser = argparse.ArgumentParser(
    description="Parse OAB (Offline Address Book) data and export to JSON or CSV.",
    formatter_class=argparse.RawTextHelpFormatter
  )

  parser.add_argument("input_file", help="Path to the input OAB file.")
  parser.add_argument(
    "-o",
    "--output_file",
    help="Path to the output file (optional). If not specified, output will be printed to stdout.",
  )
  parser.add_argument(
    "-f",
    "--format",
    choices=["json", "csv"],
    default="json",
    help="Output format: json (default) or csv.",
  )

  args = parser.parse_args()

  try:
    with open(args.input_file, "rb") as f:
      oab_data = f.read()
  except FileNotFoundError:
    print(f"Error: Input file not found: {args.input_file}", file=sys.stderr)
    sys.exit(1)
  except Exception as e:
    print(f"Error reading input file: {e}", file=sys.stderr)
    sys.exit(1)

  try:
    # Parse the OAB data
    parser = OABParser(oab_data)
    parser.parse()

    if args.output_file:
       print(f"Parsed {len(parser.get_records())} records.", file=sys.stderr)

    # Format the output
    if args.format == "json":
      output = parser.to_json()
    elif args.format == "csv":
      output = parser.to_csv()
    else:
      print(f"Error: Invalid output format: {args.format}", file=sys.stderr)
      sys.exit(1)

    # Write the output
    if args.output_file:
      try:
        if args.format == "json":
          parser.save_json(args.output_file)
        else: #CSV
          with open(args.output_file, 'w', newline='', encoding='utf-8') as csv_file:
            csv_file.write(parser.to_csv())
      except Exception as e:
        print(f"Error writing to output file: {e}", file=sys.stderr)
        sys.exit(1)
    else:
      print(output)  # Print to stdout

  except Exception as e:
    print(f"Error during OAB parsing: {e}", file=sys.stderr)
    sys.exit(1)

if __name__ == "__main__":
  main()