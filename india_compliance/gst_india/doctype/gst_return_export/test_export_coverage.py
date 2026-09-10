# Copyright (c) 2026, Resilient Tech and contributors
# For license information, please see license.txt

"""Catch a template change before a customer does."""

from frappe.tests import IntegrationTestCase

from india_compliance.gst_india.doctype.gst_return_export.audit import document_tabs, unmapped_columns
from india_compliance.gst_india.doctype.gst_return_export.gstr_2_export import (
    EXPORTERS,
    NOT_FILLED,
)


class TestTemplateCoverage(IntegrationTestCase):
    def test_every_template_column_resolves(self):
        """A blank column must be listed in KNOWN_BLANK, not a surprise."""
        for return_type, exporter in EXPORTERS.items():
            with self.subTest(return_type):
                self.assertEqual(
                    unmapped_columns(exporter),
                    [],
                    f"{exporter.template}: run report_unmapped for the paste-ready map",
                )

    def test_registry_follows_the_workbook(self):
        """Code order follows the workbook, so the two read side by side."""
        for return_type, exporter in EXPORTERS.items():
            with self.subTest(return_type):
                self.assertEqual(list(exporter.SHEETS), document_tabs(exporter.template))

    def test_registry_lists_every_tab(self):
        """A tab left out is invisible. Skipping one must be written down."""
        for return_type, exporter in EXPORTERS.items():
            with self.subTest(return_type):
                missing = set(document_tabs(exporter.template)) - set(exporter.SHEETS)
                self.assertEqual(missing, set(), f"{exporter.template}: tabs missing from SHEETS")

    def test_not_filled_tabs_are_deliberate(self):
        """A skipped tab must be a real tab."""
        for return_type, exporter in EXPORTERS.items():
            skipped = [s for s, spec in exporter.SHEETS.items() if spec is NOT_FILLED]
            with self.subTest(return_type):
                for sheet in skipped:
                    self.assertIn(sheet, document_tabs(exporter.template))
