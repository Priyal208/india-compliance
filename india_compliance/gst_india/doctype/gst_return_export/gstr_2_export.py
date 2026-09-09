# Copyright (c) 2026, Resilient Tech and contributors
# For license information, please see license.txt

"""GSTR-2A/2B exporters. 2B rows from the sync's readers; 2A raw verbatim."""

from functools import partial
from typing import ClassVar

import frappe
from frappe.query_builder.functions import Max
from frappe.utils import flt

from india_compliance.gst_india.doctype.gst_return_export.return_adapters import (
    GSTR2AAdapter,
    GSTR2BAdapter,
)
from india_compliance.gst_india.doctype.gst_return_export.template_exporter import (
    GovReturnExporter,
    amend_text,
    as_section_dict,
    date_text,
    financial_year,
    normalize_label,
    percent_text,
    period_text,
    raw_date_text,
    raw_yes_no_text,
    reformat_date,
    spec_value,
    split_label,
    state_from_code,
    state_text,
    write_cell,
    yes_no_text,
)
from india_compliance.gst_india.doctype.gst_return_log.gst_return_log import (
    DOCTYPE as RETURN_LOG,
)
from india_compliance.gst_india.utils.gstr_utils import ReturnType
from india_compliance.gst_returns.fields.gstr2 import DocField as doc
from india_compliance.gst_returns.fields.gstr2 import RawField2a as raw2a
from india_compliance.gst_returns.fields.gstr2 import RawField2b as raw2b

# each sheet generation worded the supplier filing header differently
SUPPLIER_HEADERS = (
    "gstr-1/iff",
    "gstr-1/1a/iff",
    "gstr-1/iff/1a",
    "gstr-1/iff/gstr-1a",
    "gstr-1/iff/gstr-5",
    "gstr-1/1a/iff/gstr-5",
    "gstr-1/iff/1a/gstr-5",
    "isd gstr-6",
)


class GSTR2BExporter(GovReturnExporter):
    """2B: document sheets from canonical rows, ITC summary sheets from itcsumm."""

    adapter = GSTR2BAdapter
    template = "gstr2b_excel_template_v1.1.xlsx"
    NUMBER_FORMAT = "0.00"

    # column label -> canonical field, presenter where the cell shows it differently
    FIELDS: ClassVar[dict] = {
        "gstin of supplier": doc.SUPPLIER_GSTIN,
        "gstin of isd": doc.SUPPLIER_GSTIN,
        "gstin of eco": doc.SUPPLIER_GSTIN,
        "trade/legal name": doc.SUPPLIER_NAME,
        **{f"{head} period": (doc.SUP_RETURN_PERIOD, period_text) for head in SUPPLIER_HEADERS},
        **{f"{head} filing date": (doc.GSTR_1_FILING_DATE, date_text) for head in SUPPLIER_HEADERS},
        "invoice details | invoice number": doc.BILL_NO,
        "invoice details | invoice type": doc.SUPPLY_TYPE,
        "invoice details | invoice date": (doc.BILL_DATE, date_text),
        "invoice details | invoice value": doc.DOC_VALUE,
        "document details | document number": doc.BILL_NO,  # ECO
        "document details | document type": doc.SUPPLY_TYPE,
        "document details | document date": (doc.BILL_DATE, date_text),
        "document details | document value": doc.DOC_VALUE,
        "credit note/debit note details | note number": doc.BILL_NO,
        "credit note/debit note details | note type": doc.DOC_TYPE,
        "credit note/debit note details | note supply type": doc.SUPPLY_TYPE,
        "credit note/debit note details | note date": (doc.BILL_DATE, date_text),
        "credit note/debit note details | note value": doc.DOC_VALUE,
        "debit note details | note number": doc.BILL_NO,  # B2B-DNR (ITC reversal)
        "debit note details | note type": doc.DOC_TYPE,
        "debit note details | note supply type": doc.SUPPLY_TYPE,
        "debit note details | note date": (doc.BILL_DATE, date_text),
        "debit note details | note value": doc.DOC_VALUE,
        "place of supply": (doc.POS, state_text),
        "supply attract reverse charge": (doc.REVERSE_CHARGE, yes_no_text),
        "taxable value": doc.TAXABLE_VALUE,
        "tax amount | integrated tax": doc.IGST,
        "tax amount | central tax": doc.CGST,
        "tax amount | state/ut tax": doc.SGST,
        "tax amount | cess": doc.CESS,
        "itc availability": doc.ITC_AVAILABILITY,
        "eligibility of itc": doc.ITC_AVAILABILITY,
        "reason": doc.ITC_REASON,
        "applicable % of tax rate": (doc.DIFF_PERCENTAGE, percent_text),
        "source": doc.IRN_SOURCE,
        "irn": doc.IRN_NUMBER,
        "irn date": (doc.IRN_GEN_DATE, date_text),
        "isd document type": doc.DOC_TYPE,
        "isd document number": doc.BILL_NO,
        "isd document date": (doc.BILL_DATE, date_text),
        "input tax distribution by isd | integrated tax": doc.IGST,
        "input tax distribution by isd | central tax": doc.CGST,
        "input tax distribution by isd | state/ut tax": doc.SGST,
        "input tax distribution by isd | cess": doc.CESS,
        "port code": doc.PORT_CODE,
        "bill of entry details | number": doc.BILL_NO,
        "bill of entry details | date": (doc.BILL_DATE, date_text),
        "bill of entry details | taxable value": doc.TAXABLE_VALUE,
        "amount of tax | integrated tax": doc.IGST,
        "amount of tax | cess": doc.CESS,
        # B2B-DNRA: portal template mis-merge these headers one column left; mapped by position
        "taxable value | integrated tax": doc.TAXABLE_VALUE,
        "taxable value | central tax": doc.IGST,
        "taxable value | state/ut tax": doc.CGST,
        "taxable value | cess": doc.SGST,
        "tax amount": doc.CESS,
    }

    # 'original details' block: every variant lands on the same canonical trio
    ORIGINAL_FIELDS: ClassVar[dict] = {
        "invoice number": doc.ORIGINAL_BILL_NO,
        "invoice date": (doc.ORIGINAL_BILL_DATE, date_text),
        "note type": doc.ORIGINAL_DOC_TYPE,
        "note number": doc.ORIGINAL_BILL_NO,
        "note date": (doc.ORIGINAL_BILL_DATE, date_text),
        "isd document type": doc.ORIGINAL_DOC_TYPE,
        "document number": doc.ORIGINAL_BILL_NO,
        "document date": (doc.ORIGINAL_BILL_DATE, date_text),
    }

    # portal-echo columns the canonical row lack; read from the raw record
    VERBATIM: ClassVar[dict] = {
        "original invoice number": raw2b.ORIGINAL_INVOICE_NUMBER,
        "original invoice date": (raw2b.ORIGINAL_INVOICE_DATE, raw_date_text),
        "icegate reference date": (raw2b.ICEGATE_REF_DATE, raw_date_text),
        "type of amendment": (raw2b.AMEND_TYPE, amend_text),
        "whether itc to be reduced (taxpayer's input)": (raw2b.ITC_REDUCTION_REQUIRED, raw_yes_no_text),
        "amount declared by taxpayer for itc reduction | integrated tax": raw2b.DECLARED_IGST,
        "amount declared by taxpayer for itc reduction | central tax": raw2b.DECLARED_CGST,
        "amount declared by taxpayer for itc reduction | state/ut tax": raw2b.DECLARED_SGST,
        "amount declared by taxpayer for itc reduction | cess": raw2b.DECLARED_CESS,
        "remarks": raw2b.REMARKS,
    }

    # sheet -> (payload block, reader category, amended filter for the import sheets)
    SHEETS: ClassVar[list] = [
        ("B2B", "docdata", "B2B", None),
        ("B2BA", "docdata", "B2BA", None),
        ("B2B-CDNR", "docdata", "CDNR", None),
        ("B2B-CDNRA", "docdata", "CDNRA", None),
        ("ISD", "docdata", "ISD", None),
        ("ISDA", "docdata", "ISDA", None),
        ("ECO", "docdata", "ECOM", None),
        ("ECOA", "docdata", "ECOMA", None),
        ("IMPG", "docdata", "IMPG", False),
        ("IMPGA", "docdata", "IMPG", True),
        ("IMPGSEZ", "docdata", "IMPGSEZ", False),
        ("IMPGSEZA", "docdata", "IMPGSEZ", True),
        ("B2B (ITC Reversal)", "itcrev", "B2B", None),
        ("B2BA (ITC Reversal)", "itcrev", "B2BA", None),
        ("B2B-DNR", "itcrev", "CDNR", None),
        ("B2B-DNRA", "itcrev", "CDNRA", None),
        ("B2B(Rejected)", "rejected", "B2B", None),
        ("B2BA(Rejected)", "rejected", "B2BA", None),
        ("B2B-CDNR(Rejected)", "rejected", "CDNR", None),
        ("B2B-CDNRA(Rejected)", "rejected", "CDNRA", None),
        ("ECO(Rejected)", "rejected", "ECOM", None),
        ("ECOA(Rejected)", "rejected", "ECOMA", None),
        ("ISD(Rejected)", "rejected", "ISD", None),
        ("ISDA(Rejected)", "rejected", "ISDA", None),
    ]

    ITC_SHEET_BLOCK: ClassVar[dict] = {
        "ITC Available": "itcavl",
        "ITC not available": "itcunavl",
        "ITC Reversal": "itcrev",
        "ITC Rejected": "itcRejected",
    }
    ITC_TAX_COLUMNS: ClassVar[dict] = {"igst": 4, "cgst": 5, "sgst": 6, "cess": 7}

    # section-total row label -> itcsumm bucket
    ITC_BUCKETS: ClassVar[dict] = {
        "all other itc - supplies from registered persons other than reverse charge": "nonrevsup",
        "all other itc - supplies from registered persons other than reverse charge (ims)": "nonrevsup",
        "itc reversal on account of rule 37a": "nonrevsup",
        "inward supplies from isd": "isdsup",
        "inward supplies liable for reverse charge": "revsup",
        "import of goods": "imports",
        "others": "othersup",
    }

    # detail row label -> key inside the current bucket
    ITC_DETAILS: ClassVar[dict] = {
        "b2b - invoices": "b2b",
        "b2b - invoices (ims)": "b2b",
        "b2b - invoices (amendment)": "b2ba",
        "b2b - invoices (amendment) (ims)": "b2ba",
        "b2b - debit notes": "cdnr",
        "b2b - debit notes (ims)": "cdnr",
        "b2b - credit notes (ims)": "cdnr",
        "b2b - credit notes - ims": "cdnr",
        "b2b - debit notes (amendment)": "cdnra",
        "b2b - debit notes (amendment) (ims)": "cdnra",
        "b2b - credit notes (amendment) (ims)": "cdnra",
        "b2b - credit notes (amendment) - ims": "cdnra",
        "b2b - credit notes (reverse charge)": "cdnrrev",
        "b2b - credit notes (reverse charge)(amendment)": "cdnrarev",
        "isd - invoices": "isd",
        "isd - credit notes": "isd",
        "isd - invoices (amendment)": "isda",
        "isd - credit notes (amendment)": "isda",
        "eco - documents": "ecom",
        "eco - documents (ims)": "ecom",
        "eco - documents (amendment)": "ecoma",
        "eco - documents (amendment) (ims)": "ecoma",
        "impg - import of goods from overseas": "impg",
        "impg (amendment)": "impga",
        "impgsez - import of goods from sez": "impgsez",
        "impgsez (amendment)": "impgasez",
    }

    def section_documents(self, category, groups):
        """(canonical row, raw record) pairs. Pairing trusts payload order."""
        reader = self.adapter(self.gstin).get_handler(self.periods[-1], category)
        _get_details, docs_key, _has_items = reader.SECTIONS[category]
        if docs_key:
            records = [record for group in groups for record in group.get(docs_key) or []]
        else:
            records = list(groups)
        return list(zip(reader.get_all_transactions(groups), records, strict=True))

    def fill(self):
        docdata = self.adapter.payload_sections(self.raw)
        sources = {
            "docdata": docdata,
            "itcrev": as_section_dict(docdata.get("itcrev")),
            "rejected": as_section_dict(self.raw.get("docRejdata")),
        }
        filled = False
        for sheet, source, category, amended in self.SHEETS:
            groups = sources[source].get(category.lower()) or []
            if not groups:
                continue
            documents = self.section_documents(category, groups)
            if amended is not None:
                documents = [d for d in documents if bool(d[0].get(doc.IS_AMENDED)) == amended]
            if documents and self.render(sheet, partial(self._rows, documents)):
                filled = True

        if itcsumm := self.raw.get("itcsumm"):
            for sheet in self.ITC_SHEET_BLOCK:
                if self.excel.has_sheet(sheet):
                    self._fill_itc_sheet(sheet, itcsumm)
                    filled = True
        return filled

    def _rows(self, documents, labels):
        return [
            {label: self._cell(label, row, record) for label in labels.values()} for row, record in documents
        ]

    def _cell(self, label, row, record):
        base, is_original = split_label(label)
        if is_original:
            return spec_value(self.ORIGINAL_FIELDS.get(base), row)
        if base in self.FIELDS:
            return spec_value(self.FIELDS[base], row)
        return spec_value(self.VERBATIM.get(base), record)

    def fill_readme(self):
        if not self.excel.has_sheet("Read me"):
            return
        period = self.periods[-1]
        info = self.get_gstin_names(self.gstin)
        gendt = self.raw.get(raw2b.GENERATION_DATE)

        ws = self.excel.wb["Read me"]
        self.set_merged(ws, 4, 3, financial_year(period))
        self.set_merged(ws, 5, 3, reformat_date(period, "%m%Y", "%B"))  # full month name
        self.set_merged(ws, 6, 3, self.gstin)
        self.set_merged(ws, 7, 3, info.get("legal_name") or "")
        self.set_merged(ws, 8, 3, info.get("trade_name") or "")
        self.set_merged(ws, 9, 3, raw_date_text(gendt) if gendt else "")

    def _fill_itc_sheet(self, sheet, itcsumm):
        """Walk row labels: total row set the bucket, detail row read from it. Absent = 0."""
        block = itcsumm.get(self.ITC_SHEET_BLOCK[sheet]) or {}
        ws = self.excel.wb[sheet]
        bucket = None
        for row in range(1, ws.max_row + 1):
            heading = normalize_label(ws.cell(row, 2).value)
            if new_bucket := self.ITC_BUCKETS.get(heading):
                bucket = new_bucket
                self._write_itc_row(ws, row, block.get(bucket))
            elif bucket and (key := self.ITC_DETAILS.get(heading)):
                self._write_itc_row(ws, row, (block.get(bucket) or {}).get(key))

    def _write_itc_row(self, ws, row, values):
        for field, col in self.ITC_TAX_COLUMNS.items():
            cell = ws.cell(row=row, column=col, value=flt((values or {}).get(field)))
            cell.number_format = self.NUMBER_FORMAT


class GSTR2AExporter(GovReturnExporter):
    """2A: raw codes verbatim, item rows + per-invoice total row, names looked up."""

    adapter = GSTR2AAdapter
    template = "gstr2a_excel_template_v1.0.xlsx"

    # absent numeric = 0, like the portal
    NUMERIC_ZERO_KEYS: ClassVar[set] = {
        raw2a.TAX_RATE,
        raw2a.TAXABLE_VALUE,
        raw2a.IGST,
        raw2a.CGST,
        raw2a.SGST,
        raw2a.CESS,
        raw2a.ISD_CESS,
    }

    # label -> (level, raw key); s=supplier, i=document, it=item
    RAW_FIELDS: ClassVar[dict] = {
        "gstin of supplier": ("s", raw2a.SUPPLIER_GSTIN),
        "gstin of isd": ("s", raw2a.SUPPLIER_GSTIN),
        "gstin of eco": ("s", raw2a.SUPPLIER_GSTIN),
        # filing headers per sheet wording (ECO drop the "/5")
        **{
            f"{head} filing {part}": spec
            for head in ("gstr-1/iff/gstr-1a/5", "gstr-1/iff/gstr-1a")
            for part, spec in (
                ("status", ("s", raw2a.GSTR_1_FILING_STATUS)),
                ("date", ("s", raw2a.GSTR_1_FILING_DATE)),
                ("period", ("s", raw2a.SUP_RETURN_PERIOD)),
            )
        },
        "isd gstr-6 filing status": ("s", raw2a.GSTR_1_FILING_STATUS),
        "gstr-3b filing status": ("s", raw2a.GSTR_3B_FILED),
        "effective date of cancellation": ("s", raw2a.CANCEL_DATE),
        "document details | document number": ("i", raw2a.DOC_NUMBER),
        "document details | document type": ("i", raw2a.INVOICE_TYPE),
        "document details | document date": ("i", raw2a.DOC_DATE),
        "document details | document value": ("i", raw2a.DOC_VALUE),
        "invoice details | invoice number": ("i", raw2a.DOC_NUMBER),
        "invoice details | invoice type": ("i", raw2a.INVOICE_TYPE),
        "invoice details | invoice date": ("i", raw2a.DOC_DATE),
        "invoice details | invoice value": ("i", raw2a.DOC_VALUE),
        "credit note/debit note details | note type": ("i", raw2a.NOTE_TYPE),
        "credit note/debit note details | note number": ("i", raw2a.NOTE_NUMBER),
        "credit note/debit note details | note supply type": ("i", raw2a.INVOICE_TYPE),
        "credit note/debit note details | note date": ("i", raw2a.NOTE_DATE),
        "credit note/debit note details | note value": ("i", raw2a.DOC_VALUE),
        "place of supply": ("i", raw2a.POS, state_from_code),
        "supply attract reverse charge": ("i", raw2a.REVERSE_CHARGE),
        "tax period in which amended": ("i", raw2a.OTHER_PERIOD),  # base sheets
        "original tax period in which reported": ("i", raw2a.OTHER_PERIOD),
        "tax period in which reported earlier": ("i", raw2a.OTHER_PERIOD),
        "amendment made, if any": ("i", raw2a.AMEND_TYPE),
        "source": ("i", raw2a.IRN_SOURCE),
        "irn": ("i", raw2a.IRN),
        "irn date": ("i", raw2a.IRN_DATE),
        "eligibility of itc": ("i", raw2a.ITC_ELIGIBILITY),
        "isd document type": ("i", raw2a.ISD_DOC_TYPE),
        "isd invoice number": ("i", raw2a.ISD_DOC_NUMBER),
        "isd invoice date": ("i", raw2a.ISD_DOC_DATE),
        "isd credit note number": ("i", raw2a.ISD_DOC_NUMBER),
        "isd credit note date": ("i", raw2a.ISD_DOC_DATE),
        "original invoice number": ("i", raw2a.ORIGINAL_INVOICE_NUMBER),
        "original invoice date": ("i", raw2a.ORIGINAL_INVOICE_DATE),
        "input tax distribution by isd | integrated tax": ("i", raw2a.IGST),
        "input tax distribution by isd | central tax": ("i", raw2a.CGST),
        "input tax distribution by isd | state/ut tax": ("i", raw2a.SGST),
        "input tax distribution by isd | cess": ("i", raw2a.ISD_CESS),
        "reference date (icegate)": ("i", raw2a.ICEGATE_REF_DATE),
        "port code": ("i", raw2a.PORT_CODE),
        "bill of entry details | number": ("i", raw2a.BOE_NUMBER),
        "bill of entry details | date": ("i", raw2a.BOE_DATE),
        "bill of entry details | taxable value": ("i", raw2a.TAXABLE_VALUE),
        "amount of tax | integrated tax": ("i", raw2a.IGST),
        "amount of tax | cess": ("i", raw2a.CESS),
        "amended (yes)": ("i", raw2a.IS_AMENDED),
        "amended(yes)": ("i", raw2a.IS_AMENDED),
        "rate": ("it", raw2a.TAX_RATE),
        "taxable value": ("it", raw2a.TAXABLE_VALUE),
        "tax amount | integrated tax": ("it", raw2a.IGST),
        "tax amount | central tax": ("it", raw2a.CGST),
        "tax amount | state/ut tax": ("it", raw2a.SGST),
        "tax amount | state tax": ("it", raw2a.SGST),
        "tax amount | cess": ("it", raw2a.CESS),
        "tax amount | cess amount": ("it", raw2a.CESS),
    }

    RAW_ORIGINAL: ClassVar[dict] = {
        "invoice number": ("i", raw2a.ORIGINAL_DOC_NUMBER),
        "invoice date": ("i", raw2a.ORIGINAL_DOC_DATE),
        "note type": ("i", raw2a.NOTE_TYPE),
        "note number": ("i", raw2a.ORIGINAL_NOTE_NUMBER),
        "note date": ("i", raw2a.ORIGINAL_NOTE_DATE),
    }
    _ECOA_ORIGINAL: ClassVar[dict] = {
        "document number": ("i", raw2a.ORIGINAL_DOC_NUMBER),
        "document date": ("i", raw2a.ORIGINAL_DOC_DATE),
    }
    ORIGINAL_BY_SECTION: ClassVar[dict] = {"ECOMA": _ECOA_ORIGINAL}

    RAW_FIELDS_TDS: ClassVar[dict] = {
        "gstin of deductor": ("i", raw2a.DEDUCTOR_GSTIN),
        "deductor's name": ("i", raw2a.DEDUCTOR_NAME),
        "tax period of gstr 7": ("i", raw2a.DEDUCTION_MONTH),  # MMYYYY, portal shows raw
        "taxable value": ("i", raw2a.DEDUCTED_VALUE),
        "amount of tax deducted by deductors | integrated tax": ("i", raw2a.IGST),
        "amount of tax deducted by deductors | central tax": ("i", raw2a.CGST),
        "amount of tax deducted by deductors | state/ut tax": ("i", raw2a.SGST),
    }

    RAW_FIELDS_TCS: ClassVar[dict] = {
        "gstin of e-com. operator": ("i", raw2a.ECOM_GSTIN),
        "gross value of supplies": ("i", raw2a.SUPPLY_VALUE),
        "net amount liable for tcs": ("i", raw2a.TCS_TAXABLE_VALUE),
        "total tcs amount | integrated tax": ("i", raw2a.IGST),
        "total tcs amount | central tax": ("i", raw2a.CGST),
        "total tcs amount | state/ut tax": ("i", raw2a.SGST),
    }

    FIELDS_BY_SECTION: ClassVar[dict] = {"TDS": RAW_FIELDS_TDS, "TCS": RAW_FIELDS_TCS}

    # section -> (sheet, record list key, item list key); flat ones have neither
    SECTIONS: ClassVar[dict] = {
        "B2B": ("B2B", raw2a.INVOICES, raw2a.ITEMS),
        "B2BA": ("B2BA", raw2a.INVOICES, raw2a.ITEMS),
        "CDNR": ("CDNR", raw2a.NOTES, raw2a.ITEMS),
        "CDNRA": ("CDNRA", raw2a.NOTES, raw2a.ITEMS),
        "ECOM": ("ECO", raw2a.INVOICES, raw2a.ITEMS),
        "ECOMA": ("ECOA", raw2a.INVOICES, raw2a.ITEMS),
        "ISD": ("ISD", raw2a.ISD_DOCS, ""),
        "IMPG": ("IMPG", "", ""),
        "IMPGSEZ": ("IMPG SEZ", "", ""),
        "TDS": ("TDS", "", ""),
        "TCS": ("TCS", "", ""),
    }

    def fill(self):
        docdata = self.adapter.payload_sections(self.raw)
        gstins = self._payload_gstins(docdata)
        names = {**self._supplier_names(gstins), **self._registry_names(gstins)}
        period = self.periods[-1]
        filled = False
        for section, (sheet, list_key, item_key) in self.SECTIONS.items():
            groups = docdata.get(section.lower()) or []
            if section == "IMPGSEZ":
                # SEZ carry the supplier under sgstin/tdname
                groups = [
                    {
                        **g,
                        raw2a.SUPPLIER_GSTIN: g.get(raw2a.SEZ_GSTIN),
                        raw2a.SUPPLIER_NAME: g.get(raw2a.SEZ_TRADE_NAME),
                    }
                    for g in groups
                ]
            build_rows = partial(self._build_rows, section, groups, list_key, item_key, names, period)
            if groups and self.render(sheet, build_rows):
                filled = True
        return filled

    def fill_readme(self):
        """2A Read me header: plain cells, period as MMYYYY."""
        if not self.excel.has_sheet("Read me"):
            return
        period = self.periods[-1]
        info = self.get_gstin_names(self.gstin)
        # 2A raw has no generation date; use the sync day
        synced_on = frappe.db.get_value(RETURN_LOG, f"{self.return_type}-{period}-{self.gstin}", "modified")

        ws = self.excel.wb["Read me"]
        write_cell(ws, 2, 3, self.gstin)  # C2  Taxpayer's GSTIN
        write_cell(ws, 3, 3, info.get("legal_name") or "")  # C3  Legal name
        write_cell(ws, 4, 3, info.get("trade_name") or "")  # C4  Trade name
        write_cell(ws, 2, 5, period)  # E2  Tax period (MMYYYY)
        write_cell(ws, 3, 5, financial_year(period))  # E3  Financial year
        write_cell(ws, 4, 5, synced_on.strftime("%d-%m-%Y") if synced_on else "")  # E4  Date of generation

    def _build_rows(self, section, groups, list_key, item_key, names, period, labels):
        fields = self.FIELDS_BY_SECTION.get(section, self.RAW_FIELDS)
        original = self.ORIGINAL_BY_SECTION.get(section, self.RAW_ORIGINAL)

        def cell(label, supplier, record, item):
            if self._is_trade_name_label(label):
                return supplier.get(raw2a.SUPPLIER_NAME) or names.get(supplier.get(raw2a.SUPPLIER_GSTIN))
            if label == "e-com. operator's name":
                return names.get(record.get(raw2a.ECOM_GSTIN))
            if label == "tax period of gstr 8":
                return period
            return self._raw_value(label, {"s": supplier, "i": record, "it": item}, fields, original)

        def row_for(supplier, record, item):
            return {label: cell(label, supplier, record, item) for label in labels.values()}

        rows = []
        for supplier in groups:
            for record in (supplier.get(list_key) or []) if list_key else [supplier]:
                if not item_key:
                    rows.append(row_for(supplier, record, {}))
                    continue

                items = [entry.get(raw2a.ITEM_DETAILS, entry) for entry in record.get(item_key) or []]
                rows.extend(row_for(supplier, record, item) for item in items)
                rows.append(self._total_row(row_for(supplier, record, self._item_totals(items)), labels))
                rows.append({})
        return rows

    NUMBER_LABELS = ("invoice number", "note number", "document number")

    @classmethod
    def _total_row(cls, row, labels):
        """Per-invoice total: rate blank, "-Total" on the number. Match labels without their block."""
        for label in labels.values():
            base, is_original = split_label(label)
            if base == "rate":
                row[label] = "-"
            elif not is_original and base.endswith(cls.NUMBER_LABELS) and row.get(label) is not None:
                row[label] = f"{row[label]}-Total"
        return row

    def _raw_value(self, label, containers, fields, original):
        """Cells the map can't say; rest resolve from the maps."""
        base, is_original = split_label(label)
        record = containers["i"]
        if base.startswith(("isd invoice", "isd credit note")):
            # invoice vs credit-note columns: only the matching pair fill
            if (record.get(raw2a.ISD_DOC_TYPE) == "ISDCN") != ("credit note" in base):
                return ""
        if base == "value of supplies returned":
            return round(flt(record.get(raw2a.SUPPLY_VALUE)) - flt(record.get(raw2a.TCS_TAXABLE_VALUE)), 2)

        spec = (original if is_original else fields).get(base)
        if not spec:
            return None
        level, key, *presenter = spec
        value = containers.get(level, {}).get(key)
        if value is None and key in self.NUMERIC_ZERO_KEYS:
            value = 0
        return presenter[0](value) if presenter else value

    @staticmethod
    def _is_trade_name_label(label):
        return split_label(label)[0].startswith("trade/legal name")

    @classmethod
    def _item_totals(cls, items):
        return {field: sum(flt(item.get(field)) for item in items) for field in cls.NUMERIC_ZERO_KEYS}

    def _payload_gstins(self, docdata):
        """Every GSTIN a sheet may need a name for."""
        return {
            gstin
            for section in self.SECTIONS
            for group in (docdata.get(section.lower()) or [])
            if isinstance(group, dict)
            for gstin in (
                group.get(raw2a.SUPPLIER_GSTIN),
                group.get(raw2a.SEZ_GSTIN),
                group.get(raw2a.ECOM_GSTIN),
            )
            if gstin
        }

    def _supplier_names(self, gstins):
        """Names off inward supplies, newest row wins. Only the 2B sync saves them, so no period filter."""
        if not gstins:
            return {}

        GIS = frappe.qb.DocType("GST Inward Supply")
        rows = (
            frappe.qb.from_(GIS)
            .select(GIS.supplier_gstin, GIS.supplier_name)
            .where(
                (GIS.company_gstin == self.gstin)
                & GIS.supplier_gstin.isin(list(gstins))
                & GIS.supplier_name.isnotnull()
                & (GIS.supplier_name != "")
            )
            .groupby(GIS.supplier_gstin, GIS.supplier_name)
            .orderby(Max(GIS.modified))
            .run(as_dict=True)
        )
        return {row.supplier_gstin: row.supplier_name for row in rows}

    def _registry_names(self, gstins):
        """Cached names for every payload supplier, one query."""
        if not gstins:
            return {}

        # the portal's 2A shows the legal name (its 2B the trade name)
        return {
            row.name: row.legal_name or row.trade_name
            for row in frappe.get_all(
                "GSTIN",
                filters={"name": ("in", list(gstins))},
                fields=["name", "legal_name", "trade_name"],
            )
            if row.legal_name or row.trade_name
        }


EXPORTERS = {
    ReturnType.GSTR2A.value: GSTR2AExporter,
    ReturnType.GSTR2B.value: GSTR2BExporter,
}
