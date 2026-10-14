"""python -m dataextraction.web  -> chay web app local (mac dinh 127.0.0.1:8000)."""
import argparse

import uvicorn


def main():
    parser = argparse.ArgumentParser(description="Chay web app trich xuat danh sach tai khoan")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--reload", action="store_true", help="Tu tai lai khi sua code (dev)")
    args = parser.parse_args()
    uvicorn.run("dataextraction.web.app:app", host=args.host, port=args.port, reload=args.reload)


if __name__ == "__main__":
    main()
