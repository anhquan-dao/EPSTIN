from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
MODULE_PATH = Path(__file__).resolve().parents[1] / "scripts" / "verify_excel_inputs.py"
SPEC = importlib.util.spec_from_file_location("verify_excel_inputs", MODULE_PATH)
VERIFY = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(VERIFY)


class VerifyInputTests(unittest.TestCase):
    def test_parses_dbc_and_pm_hashes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            dbc = Path(directory) / "signals.dbc"
            dbc.write_text("BO_ 1 Frame: 8 Node\n SG_ DoorLockStatus : 0|1@1+ (1,0) [0|1] \"\" Node\n")
            pm = Path(directory) / "functions.pm"
            pm.write_text("my %functions = ('Remote Lock' => 1, EngineStart => 1);")
            a2l = Path(directory) / "functions.a2l"
            a2l.write_text("/begin FUNCTION RemoteLock /end FUNCTION\n/begin MEASUREMENT EngineRunning /end MEASUREMENT")
            self.assertEqual(VERIFY.parse_signals(dbc), {"DoorLockStatus"})
            self.assertEqual(VERIFY.parse_functions(pm), {"Remote Lock", "EngineStart"})
            self.assertEqual(VERIFY.parse_functions(a2l), {"RemoteLock", "EngineRunning"})

    def test_rejects_unknown_signal_and_function(self) -> None:
        config = {
            "source_navigators": {"Variant": {"sheet_name": "Variant", "start_row": 4}, "Signals": {"sheet_name": "Signals", "start_row": 4}},
            "test_set": {"invalid_values": {"navigator": {"column_number": 5}, "delimiter": ","}, "preconditions": {"navigator": {"column_number": 4}, "assignment_delimiter": ";", "signal_value_delimiter": ","}},
            "test_type": {"navigator": {"column_number": 11}, "input_mappings": {"Signal": {"template_marker": "signal_test"}}},
            "net_logic": {"function_test_outcome": {"variant_code": {"column_number": 2}, "signal_function": {"column_number": 8}, "non_affected_variant": {"column_number": 9}, "disabled_function_assignments": {"list_delimiter": ","}}},
            "template_keywords": {"XXXCAN_SIGNALXXX": {"navigator": {"sheet": "Signals", "column_number": 3, "start_row": 4}}},
        }
        workbook = {"Variant": [{"__row_number__": "4", "__column_1__": "VAR-1", "__column_2__": "VARIANT_1"}], "Signals": [{"__row_number__": "4", "__column_3__": "UnknownSignal", "__column_4__": "Other,1", "__column_5__": "3", "__column_8__": "MissingFunction", "__column_9__": "VARIANT_1", "__column_11__": "Signal"}]}
        errors = VERIFY.verify_records(workbook, config, {"KnownSignal"}, {"KnownFunction"})
        self.assertTrue(any("UnknownSignal" in error for error in errors))
        self.assertTrue(any("MissingFunction" in error for error in errors))


if __name__ == "__main__":
    unittest.main()
