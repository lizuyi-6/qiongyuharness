import re
import unittest
from pathlib import Path

from server import gateway

ROOT = Path(__file__).resolve().parents[1]


class OutputLimitTest(unittest.TestCase):
    def test_known_models_use_official_output_ceiling(self):
        self.assertEqual(gateway.output_limit_fields("glm-5.3"), {"max_tokens": 131072})
        self.assertEqual(gateway.output_limit_fields("step-5-preview"), {"max_tokens": 64000})
        self.assertEqual(gateway.output_limit_fields("deepseek-v4.1-flash-1017"), {"max_tokens": 393216})
        self.assertEqual(gateway.output_limit_fields("k3.1-pro"), {"max_tokens": 131072})

    def test_unknown_models_send_no_max_tokens(self):
        self.assertEqual(gateway.output_limit_fields("deepseek-v4-flash"), {})
        self.assertEqual(gateway.output_limit_fields("stepaudio-2.5-chat"), {})

    def test_no_fixed_token_cap_left(self):
        self.assertFalse(hasattr(gateway, "TOKEN_LIMIT"))

    def test_desktop_gateway_uses_the_same_table(self):
        source = (ROOT / "desktop" / "lib" / "loopback-gateway.cjs").read_text(encoding="utf-8")
        block = re.search(r"MODEL_OUTPUT_LIMITS = Object\.freeze\(\{(.*?)\}\);", source, re.S).group(1)
        desktop = {name: int(value.replace("_", "")) for name, value in re.findall(r'"([^"]+)":\s*([\d_]+)', block)}
        self.assertEqual(desktop, gateway.MODEL_OUTPUT_LIMITS)


if __name__ == "__main__":
    unittest.main()
