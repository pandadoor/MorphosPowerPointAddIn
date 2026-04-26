import argparse
import os
import shutil
import tempfile
import zipfile


REPLACEMENTS = (
    ("Arial", "A Morphos Missing Verification Font"),
    ("Calibri", "B Morphos Missing Verification Font"),
)


def patch_xml_file(path):
    with open(path, "r", encoding="utf-8") as handle:
        content = handle.read()

    updated = content
    replacement_count = 0
    for original, replacement in REPLACEMENTS:
        original_token = f'typeface="{original}"'
        replacement_token = f'typeface="{replacement}"'
        if original_token in updated:
            updated = updated.replace(original_token, replacement_token, 1)
            replacement_count += 1

    if replacement_count > 0:
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(updated)

    return replacement_count


def create_fixture(source_path, output_path):
    temp_dir = tempfile.mkdtemp(prefix="morphos-missing-font-fixture-")
    try:
        with zipfile.ZipFile(source_path, "r") as archive:
            archive.extractall(temp_dir)

        xml_root = temp_dir
        total_replacements = 0
        for current_root, _, files in os.walk(xml_root):
            for file_name in files:
                if not file_name.endswith(".xml"):
                    continue

                total_replacements += patch_xml_file(os.path.join(current_root, file_name))

        if total_replacements < len(REPLACEMENTS):
            raise RuntimeError(
                f"Expected to replace at least {len(REPLACEMENTS)} font references, but only replaced {total_replacements}."
            )

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for current_root, _, files in os.walk(temp_dir):
                for file_name in files:
                    full_path = os.path.join(current_root, file_name)
                    relative_path = os.path.relpath(full_path, temp_dir)
                    archive.write(full_path, relative_path)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    source_path = os.path.abspath(args.source)
    output_path = os.path.abspath(args.output)
    if not os.path.exists(source_path):
        raise FileNotFoundError(source_path)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    create_fixture(source_path, output_path)
    print(output_path, flush=True)


if __name__ == "__main__":
    main()
