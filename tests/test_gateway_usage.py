import unittest

from server.gateway import usage_from


class GatewayUsageTest(unittest.TestCase):
    def test_desktop_output_metadata_is_cleaned_and_preserved(self):
        usage = usage_from({
            "prompt_tokens": 8,
            "max_tokens_clamped": [{"requested": 5000, "applied": 3500}],
            "finish_reasons": ["length", "stop"],
            "truncated": True,
        })
        self.assertEqual(usage, {
            "prompt_tokens": 8,
            "max_tokens_clamped": [{"requested": 5000, "applied": 3500}],
            "finish_reasons": ["length", "stop"],
            "truncated": True,
        })

    def test_invalid_output_metadata_is_dropped(self):
        usage = usage_from({
            "prompt_tokens": True,
            "max_tokens_clamped": [{"requested": True, "applied": 3500}],
            "finish_reasons": ["length", 42],
            "truncated": "true",
        })
        self.assertEqual(usage, {})

    def test_output_metadata_lists_over_four_items_are_dropped(self):
        usage = usage_from({
            "max_tokens_clamped": [{"requested": 5000 + i, "applied": 3500} for i in range(5)],
            "finish_reasons": ["length"] * 5,
        })
        self.assertEqual(usage, {})


if __name__ == "__main__":
    unittest.main()
