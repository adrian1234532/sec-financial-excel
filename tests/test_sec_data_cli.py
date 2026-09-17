import json
import tempfile
import unittest
from pathlib import Path

from sec_data_cli import build_parser, cli_main


class FakeEngine:
    def __init__(self):
        self.calls = []
        self.identity_reset = False

    def _reset_cached_sec_identity(self):
        self.identity_reset = True

    def _normalize_ticker_queue(self, tickers):
        return [str(item).upper() for item in tickers]

    def run_ticker_queue(self, tickers, limit, **options):
        self.calls.append((list(tickers), limit, options))
        return [{"ticker": tickers[0], "status": "saved"}]

    def main(self, ticker, limit, **options):
        self.calls.append((ticker, limit, options))
        return f"output/{ticker}.xlsx"

    def _write_queue_child_result(self, path, payload):
        if path:
            Path(path).write_text(json.dumps(payload), encoding="utf-8")

    def _record_ticker_failure(self, *args, **kwargs):
        return None


class SecDataCliTests(unittest.TestCase):
    def test_parser_keeps_documented_defaults(self):
        args = build_parser().parse_args(["--ticker", "RXRX"])
        self.assertEqual(args.ticker, ["RXRX"])
        self.assertEqual(args.limit, 50)
        self.assertFalse(args.xlsx)
        self.assertFalse(args.no_arelle)

    def test_parent_dispatches_without_embedding_cli_in_engine(self):
        engine = FakeEngine()
        code = cli_main(
            ["--ticker", "RXRX", "TLN", "--xlsx", "--no-arelle", "--limit", "12"],
            engine=engine,
        )
        self.assertEqual(code, 0)
        tickers, limit, options = engine.calls[0]
        self.assertEqual(tickers, ["RXRX", "TLN"])
        self.assertEqual(limit, 12)
        self.assertTrue(options["save_xlsx"])
        self.assertFalse(options["use_arelle"])

    def test_child_writes_machine_readable_result(self):
        engine = FakeEngine()
        with tempfile.TemporaryDirectory() as directory:
            result_path = Path(directory) / "result.json"
            code = cli_main(
                [
                    "--ticker", "rxrx", "--queue-child", "--queue-result-file",
                    str(result_path), "--xlsx",
                ],
                engine=engine,
            )
            payload = json.loads(result_path.read_text(encoding="utf-8"))
        self.assertEqual(code, 0)
        self.assertEqual(payload["ticker"], "RXRX")
        self.assertEqual(payload["status"], "saved")


if __name__ == "__main__":
    unittest.main()
