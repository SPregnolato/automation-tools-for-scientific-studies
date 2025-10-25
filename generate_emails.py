"""
generate_emails.py
------------------
Generate customized plain-text email messages from an Excel file and a Jinja2 template.

Usage:
    python generate_emails.py path/to/contacts.xlsx template.txt

Behavior:
- Reads all sheets from the Excel file.
- Renders all rows using the specified Jinja2 template (from ./templates/).
- Creates one text file per sheet in the same folder as the Excel file.
"""

import sys
from pathlib import Path
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from datetime import datetime

# === Configuration ===
TEMPLATE_DIR = Path(__file__).parent / "templates"
# ======================


def format_time(value):
    """
    Convert a time string 'HH:MM:SS' to 'HH:MM AM/PM'.
    Handles empty or invalid values gracefully.
    """
    if value is None or value == "":
        return ""

    # If already a string
    if isinstance(value, str):
        try:
            dt = datetime.strptime(value, "%H:%M:%S")
        except ValueError:
            return value  # fallback if parsing fails
    else:
        dt = value  # if datetime object already

    return dt.strftime("%I:%M %p")  # e.g., '09:00 AM'


def generate_emails(excel_path: Path, template_file: str):
    """Generate plain-text emails from all sheets in an Excel file using a specified template."""

    if not excel_path.exists():
        print(f"❌ Excel file not found: {excel_path}")
        sys.exit(1)

    template_path = TEMPLATE_DIR / template_file
    if not template_path.exists():
        print(f"❌ Template file not found: {template_path}")
        sys.exit(1)

    if not TEMPLATE_DIR.exists():
        print(f"❌ Template folder not found: {TEMPLATE_DIR}")
        sys.exit(1)

    # Load all sheets
    try:
        sheets = pd.read_excel(excel_path, sheet_name=None, dtype=str)
    except Exception as e:
        print(f"❌ Failed to read Excel file: {e}")
        sys.exit(1)

    # Load Jinja2 template
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
    template = env.get_template(template_file)

    output_dir = excel_path.parent
    total_messages = 0

    for sheet_name, df in sheets.items():
        df = df.fillna("")
        messages = []

        for _, row in df.iterrows():
            context = {col: row[col] for col in df.columns}

            # Format the 'time' field if it exists
            if "Time" in context:
                context["Time"] = format_time(context["Time"])

                text = template.render(**context).strip()

                to_line = f"To: {row.get('email', '').strip()}"
                separator = "-" * 60
                messages.append(f"{to_line}\n\n{text}\n\n{separator}\n")

        # Sanitize sheet name for filename
        safe_sheet_name = "".join(
            c if c.isalnum() or c in ("-", "_") else "_" for c in sheet_name
        )
        output_path = output_dir / f"emails_output_{safe_sheet_name}.txt"

        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(messages))

        total_messages += len(df)
        print(f"✅ {len(df)} messages written to {output_path}")

    print(
        f"\n🎉 Done! Generated {total_messages} total messages across {len(sheets)} sheet(s)."
    )


if __name__ == "__main__":
    # Read Excel path and template from CLI arguments or prompt
    if len(sys.argv) > 2:
        excel_path = Path(sys.argv[1])
        template_file = sys.argv[2]
    else:
        excel_path = Path(input("Enter path to Excel file: ").strip())
        template_file = input("Enter template file name (from ./templates/): ").strip()

    generate_emails(excel_path, template_file)
