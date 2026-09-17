import unittest

import pandas as pd

import sec_data


class SemanticFallbackTests(unittest.TestCase):
    def test_rd_breakdown_cannot_inherit_revenue_context(self):
        table = pd.DataFrame(
            [
                ["Research and development expense", None, None],
                ["Platform", "$ 42,084", "$ 80,023"],
                ["Discovery", "14,753", "22,586"],
                ["Clinical", "23,661", "19,568"],
                ["Stock based compensation", "10,934", "14,089"],
                ["UK R&D tax credit", "(1,913)", "(2,064)"],
            ],
            columns=[
                "(in thousands)",
                "Three months ended June 30, 2026",
                "Three months ended June 30, 2025",
            ],
        )
        normalized = sec_data._gb_normalize_table_grid(table, 2026)
        role, prefix, confidence, basis = (
            sec_data._gb_classify_business_table_role(
                table,
                normalized,
                "Revenue discussion immediately before the R&D table",
                "Revenue",
            )
        )

        self.assertEqual(
            role, sec_data._GB_ROLE_RESEARCH_DEVELOPMENT_EXPENSE
        )
        self.assertEqual(prefix, "Research & Development Expense")
        self.assertGreaterEqual(confidence, 0.99)
        self.assertEqual(
            basis, "research_and_development_expense_breakdown"
        )
        self.assertEqual(sec_data._gb_route_for_table_role(role), "6_Disclosures")

    def test_rd_section_header_sets_expense_prefix(self):
        self.assertEqual(
            sec_data._gb_section_from_header(
                "Research and development expense"
            ),
            (True, "Research & Development Expense"),
        )


if __name__ == "__main__":
    unittest.main()
