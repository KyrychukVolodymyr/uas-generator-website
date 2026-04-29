import argparse
import sys

from license_manager import activate_license


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--key", required=True)
    parser.add_argument("--email", required=True)
    parser.add_argument("--terms-version", default="2026-04-21-v2")
    parser.add_argument("--app-version", default="v1")
    args = parser.parse_args()
    try:
        result = activate_license(
            email=args.email.strip(),
            license_key=args.key.strip(),
            accepted_terms_version=args.terms_version.strip(),
            app_version=args.app_version.strip(),
        )
        print("ACTIVATION_OK=True")
        print(f"DEVICE_ID={result.get('device_id', '')}")
        print("LICENSE_STORED=True")
        return 0
    except Exception as exc:
        print("ACTIVATION_OK=False")
        print(f"ERROR={exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
