import re
from typing import Dict, List, Optional, Tuple, Any
import pandas as pd


def clean_text(val: Any) -> str:
    """Safely converts any cell value to a clean string."""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def parse_invoice_entry(val: Any) -> Optional[Dict[str, Any]]:
    """
    Parses an invoice text or number into its structural parts:
    - series: e.g. 'F85S', 'FP02', 'E001', 'B001'
    - raw_num: e.g. '0000064535'
    - clean_num: e.g. '64535'
    - norm_full: e.g. 'F85S-0000064535'
    - norm_short: e.g. 'F85S-64535'
    """
    s = clean_text(val)
    if not s or s.lower() == "nan":
        return None

    # Remove internal spaces around hyphen, slash, etc.
    s_compact = re.sub(r"\s+", " ", s).strip()
    s_no_spaces = re.sub(r"\s*([-_/])\s*", r"\1", s_compact)
    s_upper = s_no_spaces.upper()

    # Pattern 1: Standard Series-Separator-Number (e.g., F85S-0000064535, FP02/69670, E001 123)
    # Series: 1 to 6 alphanumeric characters starting with letter or number
    m = re.match(r"^([A-Z0-9]{1,6})[-_/ ]+([0-9]+)$", s_upper)
    if not m:
        # Pattern 2: Series and number without separator (e.g., F85S0000064535, FP020000069670)
        # where series is 4 alphanumeric (typical SUNAT series) followed by 4+ digits
        m = re.match(r"^([A-Z][A-Z0-9]{3})([0-9]{4,10})$", s_upper)

    if m:
        series = m.group(1).upper()
        raw_num = m.group(2)
        clean_num = raw_num.lstrip("0") or "0"
        return {
            "original": s,
            "series": series,
            "raw_num": raw_num,
            "clean_num": clean_num,
            "norm_full": f"{series}-{raw_num}",
            "norm_short": f"{series}-{clean_num}",
            "type": "SERIES_NUM",
        }

    # Pattern 3: Digits only (e.g. 0000064535 or 64535)
    digits_only = re.sub(r"\D", "", s_upper)
    if s_upper.isdigit():
        raw_num = s_upper
        clean_num = raw_num.lstrip("0") or "0"
        return {
            "original": s,
            "series": None,
            "raw_num": raw_num,
            "clean_num": clean_num,
            "norm_full": raw_num,
            "norm_short": clean_num,
            "type": "NUM_ONLY",
        }

    # Pattern 4: Other text / unparsed format
    return {
        "original": s,
        "series": None,
        "raw_num": digits_only if digits_only else None,
        "clean_num": (digits_only.lstrip("0") or "0") if digits_only else None,
        "norm_full": s_upper,
        "norm_short": s_upper,
        "type": "RAW_TEXT",
    }


class ExcelInvoiceIndex:
    """
    Builds an ultra-fast index over the target Excel column.
    Allows O(1) lookups for:
    - Exact match (series + raw_num)
    - Normalized match (series + clean_num)
    - Correlative with zeros (raw_num)
    - Number without zeros (clean_num)
    - Substring containment
    """

    def __init__(self, df: pd.DataFrame, col_name: str):
        self.df = df
        self.col_name = col_name

        # Mappings: key -> list of row indices (0-based in df, excel row = index + 2 if header on row 1)
        self.exact_map: Dict[str, List[int]] = {}
        self.series_clean_map: Dict[Tuple[str, str], List[int]] = {}
        self.zeros_num_map: Dict[str, List[int]] = {}
        self.clean_num_map: Dict[str, List[int]] = {}
        self.parsed_rows: Dict[int, Optional[Dict[str, Any]]] = {}

        self._build_index()

    def _build_index(self):
        for idx, val in self.df[self.col_name].items():
            parsed = parse_invoice_entry(val)
            self.parsed_rows[idx] = parsed
            if not parsed:
                continue

            # Exact full normalized string
            norm_full = parsed["norm_full"]
            self.exact_map.setdefault(norm_full, []).append(idx)

            # Raw uppercase string
            raw_orig = clean_text(val).upper()
            if raw_orig and raw_orig != norm_full:
                self.exact_map.setdefault(raw_orig, []).append(idx)

            # Series + clean_num
            if parsed["series"] and parsed["clean_num"]:
                key = (parsed["series"], parsed["clean_num"])
                self.series_clean_map.setdefault(key, []).append(idx)

            # Zeros num (e.g. '0000064535')
            if parsed["raw_num"]:
                self.zeros_num_map.setdefault(parsed["raw_num"], []).append(idx)

            # Clean num (e.g. '64535')
            if parsed["clean_num"]:
                self.clean_num_map.setdefault(parsed["clean_num"], []).append(idx)


def match_invoice(
    query_str: str,
    index: ExcelInvoiceIndex,
    match_exact: bool = True,
    match_zeros: bool = True,
    match_no_zeros: bool = True,
    match_diff_series: bool = False,
) -> List[Dict[str, Any]]:
    """
    Matches a single query string against the Excel invoice index.
    Returns a list of match result dictionaries.
    """
    q_parsed = parse_invoice_entry(query_str)
    if not q_parsed:
        return []

    results = []
    seen_row_indices = set()

    # 1. Exact match (F85S-0000064535 or exact original)
    if match_exact:
        exact_keys = [q_parsed["norm_full"]]
        orig_clean = clean_text(query_str).upper()
        if orig_clean not in exact_keys:
            exact_keys.append(orig_clean)

        for key in exact_keys:
            for r_idx in index.exact_map.get(key, []):
                if r_idx not in seen_row_indices:
                    seen_row_indices.add(r_idx)
                    results.append(
                        {
                            "row_index": r_idx,
                            "match_type": "Exacta (Serie y Correlativo idéntico)",
                            "confidence": "Alta",
                            "level": 1,
                        }
                    )

        # 1b. Same Series + Same clean number (e.g. F85S-0000064535 vs F85S-64535)
        if q_parsed["series"] and q_parsed["clean_num"]:
            pair = (q_parsed["series"], q_parsed["clean_num"])
            for r_idx in index.series_clean_map.get(pair, []):
                if r_idx not in seen_row_indices:
                    seen_row_indices.add(r_idx)
                    cell_parsed = index.parsed_rows[r_idx]
                    raw_c = cell_parsed["raw_num"] if cell_parsed else ""
                    results.append(
                        {
                            "row_index": r_idx,
                            "match_type": f"Serie idéntica ({q_parsed['series']}) con distinto formato de ceros ({q_parsed['raw_num']} vs {raw_c})",
                            "confidence": "Alta",
                            "level": 2,
                        }
                    )

    # 2. Match with zeros (e.g. 0000064535)
    if match_zeros and q_parsed["raw_num"]:
        raw_num = q_parsed["raw_num"]
        for r_idx in index.zeros_num_map.get(raw_num, []):
            if r_idx not in seen_row_indices:
                cell_parsed = index.parsed_rows[r_idx]
                cell_series = cell_parsed["series"] if cell_parsed else None

                # If query had a series and cell has a different series
                if q_parsed["series"] and cell_series and q_parsed["series"] != cell_series:
                    if match_diff_series:
                        seen_row_indices.add(r_idx)
                        results.append(
                            {
                                "row_index": r_idx,
                                "match_type": f"Mismo correlativo con ceros ({raw_num}) pero serie distinta ({q_parsed['series']} vs {cell_series})",
                                "confidence": "Media (Verificar serie)",
                                "level": 4,
                            }
                        )
                else:
                    seen_row_indices.add(r_idx)
                    detail = f" [Serie Excel: {cell_series}]" if cell_series else ""
                    results.append(
                        {
                            "row_index": r_idx,
                            "match_type": f"Correlativo con ceros ({raw_num}){detail}",
                            "confidence": "Media",
                            "level": 3,
                        }
                    )

    # 3. Match without zeros (e.g. 64535)
    if match_no_zeros and q_parsed["clean_num"]:
        clean_num = q_parsed["clean_num"]
        for r_idx in index.clean_num_map.get(clean_num, []):
            if r_idx not in seen_row_indices:
                cell_parsed = index.parsed_rows[r_idx]
                cell_series = cell_parsed["series"] if cell_parsed else None

                # If query had a series and cell has a different series
                if q_parsed["series"] and cell_series and q_parsed["series"] != cell_series:
                    if match_diff_series:
                        seen_row_indices.add(r_idx)
                        results.append(
                            {
                                "row_index": r_idx,
                                "match_type": f"Mismo número sin ceros ({clean_num}) pero serie distinta ({q_parsed['series']} vs {cell_series})",
                                "confidence": "Baja (Verificar serie)",
                                "level": 5,
                            }
                        )
                else:
                    seen_row_indices.add(r_idx)
                    detail = f" [Serie Excel: {cell_series}]" if cell_series else ""
                    results.append(
                        {
                            "row_index": r_idx,
                            "match_type": f"Número sin ceros ({clean_num}){detail}",
                            "confidence": "Media",
                            "level": 3,
                        }
                    )

    return results


def check_invoices_batch(
    df: pd.DataFrame,
    col_name: str,
    raw_query_text: str,
    match_exact: bool = True,
    match_zeros: bool = True,
    match_no_zeros: bool = True,
    match_diff_series: bool = False,
    header_row_offset: int = 2,
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    """
    Processes all invoice lines pasted by the user.
    Returns:
    - summary_df: one row per pasted invoice with status, match count, best match type, etc.
    - details_df: one row per matching excel row found.
    - stats: summary metrics dictionary.
    """
    lines = [line.strip() for line in raw_query_text.strip().splitlines() if line.strip()]
    if not lines:
        return pd.DataFrame(), pd.DataFrame(), {}

    index = ExcelInvoiceIndex(df, col_name)

    summary_records = []
    detail_records = []

    # Count occurrences in query list to detect internal duplicates
    query_counts = {}
    for line in lines:
        query_counts[line] = query_counts.get(line, 0) + 1

    total_queries = len(lines)
    found_count = 0
    not_found_count = 0
    multiple_matches_count = 0

    for query_pos, query_line in enumerate(lines, start=1):
        parsed_q = parse_invoice_entry(query_line)
        matches = match_invoice(
            query_line,
            index,
            match_exact=match_exact,
            match_zeros=match_zeros,
            match_no_zeros=match_no_zeros,
            match_diff_series=match_diff_series,
        )

        internal_dup = query_counts[query_line] > 1

        if matches:
            found_count += 1
            if len(matches) > 1:
                multiple_matches_count += 1
            status = "DUPLICADO" if len(matches) == 1 else f"DUPLICADO ({len(matches)} veces)"
            # Sort by level (highest confidence first)
            matches.sort(key=lambda x: x["level"])
            best_match_type = matches[0]["match_type"]
            best_confidence = matches[0]["confidence"]
        else:
            not_found_count += 1
            status = "NO ENCONTRADO"
            best_match_type = "-"
            best_confidence = "-"

        summary_records.append(
            {
                "N°": query_pos,
                "Texto Consultado": query_line,
                "Estado en Excel": status,
                "Total Coincidencias": len(matches),
                "Mejor Coincidencia": best_match_type,
                "Nivel Confianza": best_confidence,
                "Repetido en Lista": "SÍ (Interno)" if internal_dup else "No",
                "Serie Detectada": parsed_q["series"] if parsed_q else "-",
                "Correlativo Ceros": parsed_q["raw_num"] if parsed_q else "-",
                "Número Sin Ceros": parsed_q["clean_num"] if parsed_q else "-",
            }
        )

        for match in matches:
            r_idx = match["row_index"]
            excel_row_num = r_idx + header_row_offset  # Excel 1-based row number
            excel_val = df.at[r_idx, col_name]

            # Build record with Excel row data
            rec = {
                "Item Consultado": query_line,
                "Fila Excel": excel_row_num,
                "Tipo Coincidencia": match["match_type"],
                "Confianza": match["confidence"],
                f"Excel [{col_name}]": excel_val,
            }

            # Add other columns from the Excel row for context
            for c in df.columns:
                if c != col_name:
                    rec[str(c)] = df.at[r_idx, c]

            detail_records.append(rec)

    summary_df = pd.DataFrame(summary_records)
    details_df = pd.DataFrame(detail_records)

    stats = {
        "total_queries": total_queries,
        "unique_queries": len(query_counts),
        "found_count": found_count,
        "not_found_count": not_found_count,
        "multiple_matches_count": multiple_matches_count,
        "internal_duplicates": sum(1 for v in query_counts.values() if v > 1),
    }

    return summary_df, details_df, stats
