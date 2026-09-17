"""Command-line entry point for the SEC financial workbook engine.

The extraction and accounting implementation remains in :mod:`sec_data`.
Keeping argument parsing and process orchestration here makes the engine
importable without coupling it to the command-line interface.
"""

from __future__ import annotations

import argparse
import importlib
import sys
import traceback
from typing import Any, Sequence


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--ticker",
        required=True,
        nargs="+",
        metavar="TICKER",
        help=(
            "One or more ticker symbols. Multiple symbols run as isolated "
            "sequential child processes, for example: --ticker AMZN GOOGL UBER"
        ),
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=50,
        help="Number of filings to pull per ticker (fewer = faster; less history).",
    )
    parser.add_argument(
        "--no-arelle",
        action="store_true",
        help=(
            "Skip the slow Arelle custom-tag pre-pass "
            "(10-K annual filings and 20-F/40-F annual enrichment)."
        ),
    )
    parser.add_argument("--log", action="store_true", help="Print detailed logs")
    parser.add_argument(
        "--xlsx",
        action="store_true",
        help="Save as .xlsx instead of CSV.",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        help="Native 10-K/10-Q worker count inside each isolated ticker process.",
    )
    parser.add_argument(
        "--annual",
        action="store_true",
        help="For native filers, fetch annual filings only and output FY columns.",
    )
    parser.add_argument(
        "--quality",
        "--save-quality",
        dest="save_quality",
        action="store_true",
        help="Also save fact-audit and operating-alias CSV files under output/quality.",
    )
    parser.add_argument(
        "--stop-on-error",
        action="store_true",
        help="Stop a multi-ticker queue when one ticker fails.",
    )
    parser.add_argument(
        "--reset-identity",
        action="store_true",
        help="Delete the saved SEC contact identity and show first-run setup again.",
    )
    parser.add_argument("--queue-child", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--queue-result-file", default=None, help=argparse.SUPPRESS)
    return parser


def _load_engine() -> Any:
    # The legacy engine is intentionally loaded at runtime. This keeps the
    # small typed CLI boundary independent from the gradually typed monolith.
    return importlib.import_module("sec_data")


def cli_main(argv: Sequence[str] | None = None, *, engine: Any = None) -> int:
    """Parse CLI arguments and dispatch to the financial engine."""
    parser = build_parser()
    args = parser.parse_args(argv)
    engine = engine or _load_engine()

    if args.reset_identity:
        engine._reset_cached_sec_identity()

    if args.queue_child:
        child_tickers = engine._normalize_ticker_queue(args.ticker)
        if len(child_tickers) != 1:
            parser.error("A queue child must receive exactly one ticker.")
        ticker = child_tickers[0]
        try:
            output_path = engine.main(
                ticker,
                args.limit,
                use_arelle=not args.no_arelle,
                log_output=args.log,
                save_xlsx=args.xlsx,
                workers=args.workers,
                annual=args.annual,
                save_quality=args.save_quality,
            )
            engine._write_queue_child_result(
                args.queue_result_file,
                {
                    "ticker": ticker,
                    "status": "saved" if output_path else "no_data",
                    "output_path": output_path,
                    "error": None,
                },
            )
            return 0
        except KeyboardInterrupt:
            raise
        except Exception as exc:
            trace_text = traceback.format_exc()
            error_log = engine._record_ticker_failure(
                ticker,
                exc,
                traceback_text=trace_text,
                emit_to_stderr=False,
                context={
                    "mode": "queue_child",
                    "limit": args.limit,
                    "workers": args.workers,
                    "annual": args.annual,
                    "arelle": not args.no_arelle,
                },
            )
            if error_log:
                print(
                    f"[ERROR] Full traceback saved to: {error_log}",
                    file=sys.stderr,
                    flush=True,
                )
            engine._write_queue_child_result(
                args.queue_result_file,
                {
                    "ticker": ticker,
                    "status": "failed",
                    "output_path": None,
                    "error": f"{type(exc).__name__}: {exc}",
                    "error_log": error_log,
                },
            )
            return 1

    try:
        results = engine.run_ticker_queue(
            args.ticker,
            args.limit,
            use_arelle=not args.no_arelle,
            log_output=args.log,
            save_xlsx=args.xlsx,
            workers=args.workers,
            annual=args.annual,
            stop_on_error=args.stop_on_error,
            save_quality=args.save_quality,
        )
    except ValueError as exc:
        parser.error(str(exc))
    return 1 if any(result["status"] == "failed" for result in results) else 0


if __name__ == "__main__":
    raise SystemExit(cli_main())
