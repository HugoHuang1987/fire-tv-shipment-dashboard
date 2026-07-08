from __future__ import annotations

import argparse
import hashlib
import json
import re
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path


CANONICAL_HEADERS = {
    "customer": {"客户", "客户类别"},
    "category": {"大类", "产品平台", "平台"},
    "project": {"项目", "项目名称"},
    "chip": {"机芯", "机芯方案"},
    "region": {"区域", "市场区域"},
    "engine": {"内部机芯编码"},
    "bomRegion": {"bom区域", "bom地区"},
    "cooperation": {"合作模式"},
    "osVersion": {"系统版本", "os版本"},
}


def clean_text(value: str) -> str:
    value = value.replace("\u200b", "").replace("\xa0", " ")
    lines = [re.sub(r"\s+", " ", line).strip() for line in value.splitlines()]
    return "\n".join(line for line in lines if line)


def normalize_header(value: str) -> str:
    return re.sub(r"[\s_\-:：/（）()]", "", clean_text(value)).lower()


class TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[dict]]] = []
        self.table_depth = 0
        self.current_table: list[list[dict]] | None = None
        self.current_row: list[dict] | None = None
        self.current_cell: dict | None = None

    def handle_starttag(self, tag: str, attrs) -> None:
        attrs_dict = dict(attrs)
        if tag == "table":
            self.table_depth += 1
            if self.table_depth == 1:
                self.current_table = []
        elif self.table_depth == 1 and tag == "tr":
            self.current_row = []
        elif self.table_depth == 1 and tag in {"th", "td"} and self.current_row is not None:
            self.current_cell = {
                "text": [],
                "rowspan": max(1, int(attrs_dict.get("rowspan", "1") or "1")),
                "colspan": max(1, int(attrs_dict.get("colspan", "1") or "1")),
            }
        elif self.current_cell is not None and tag == "br":
            self.current_cell["text"].append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in {"th", "td"} and self.current_cell is not None and self.current_row is not None:
            self.current_cell["text"] = clean_text("".join(self.current_cell["text"]))
            self.current_row.append(self.current_cell)
            self.current_cell = None
        elif tag == "tr" and self.current_row is not None and self.current_table is not None:
            if self.current_row:
                self.current_table.append(self.current_row)
            self.current_row = None
        elif tag == "table":
            if self.table_depth == 1 and self.current_table:
                self.tables.append(self.current_table)
                self.current_table = None
            self.table_depth = max(0, self.table_depth - 1)

    def handle_data(self, data: str) -> None:
        if self.current_cell is not None:
            self.current_cell["text"].append(data)


def expand_table(raw_rows: list[list[dict]]) -> list[list[str]]:
    active: dict[int, tuple[str, int]] = {}
    expanded: list[list[str]] = []
    for raw_row in raw_rows:
        row_values: dict[int, str] = {column: value for column, (value, _) in active.items()}
        next_active: dict[int, tuple[str, int]] = {
            column: (value, remaining - 1)
            for column, (value, remaining) in active.items()
            if remaining > 1
        }
        column = 0
        for cell in raw_row:
            while column in row_values:
                column += 1
            text = cell["text"]
            for offset in range(cell["colspan"]):
                target = column + offset
                row_values[target] = text
                if cell["rowspan"] > 1:
                    next_active[target] = (text, cell["rowspan"] - 1)
            column += cell["colspan"]
        active = next_active
        if row_values:
            width = max(row_values) + 1
            expanded.append([row_values.get(index, "") for index in range(width)])
    return expanded


def identify_columns(header: list[str]) -> dict[str, int]:
    normalized = [normalize_header(value) for value in header]
    mapping: dict[str, int] = {}
    for canonical, aliases in CANONICAL_HEADERS.items():
        alias_set = {normalize_header(alias) for alias in aliases}
        for index, value in enumerate(normalized):
            if value in alias_set:
                mapping[canonical] = index
                break
    return mapping


def extract_payload(html_path: Path, source_url: str) -> dict:
    parser = TableParser()
    html_bytes = html_path.read_bytes()
    html_text = html_bytes.decode("utf-8", errors="replace")
    parser.feed(html_text)
    candidates = []
    for table_index, raw_table in enumerate(parser.tables):
        table = expand_table(raw_table)
        for header_index, row in enumerate(table):
            mapping = identify_columns(row)
            if "engine" in mapping:
                candidates.append((len(table) - header_index, table_index, header_index, table, mapping))
    if not candidates:
        raise ValueError("页面中未找到包含“内部机芯编码”的表格；可能需要登录，或Wiki页面结构已变化。")

    _, table_index, header_index, table, mapping = max(candidates)
    required_columns = {"customer", "category", "project", "chip", "region", "engine"}
    missing_columns = sorted(required_columns - set(mapping))
    if missing_columns:
        raise ValueError(f"Wiki表格缺少必要列：{', '.join(missing_columns)}；规则表未修改。")

    rows = []
    context = {key: "" for key in ("customer", "category", "project", "chip", "region")}
    for source_row, row in enumerate(table[header_index + 1 :], start=header_index + 2):
        record = {}
        for key in CANONICAL_HEADERS:
            index = mapping.get(key)
            record[key] = clean_text(row[index]) if index is not None and index < len(row) else ""
        for key in context:
            if record[key]:
                context[key] = record[key]
            else:
                record[key] = context[key]
        if record["engine"] or any(record[key] for key in context):
            record["sourceRow"] = source_row
            rows.append(record)

    if not any(record["engine"] for record in rows):
        raise ValueError("Wiki表格已找到，但“内部机芯编码”列没有有效内容。")
    return {
        "schemaVersion": 2,
        "sourceUrl": source_url,
        "extractedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
        "sourceSha256": hashlib.sha256(html_bytes).hexdigest(),
        "tableIndex": table_index,
        "headerSourceRow": header_index + 1,
        "tableDataRowCount": max(0, len(table) - header_index - 1),
        "headerRow": table[header_index],
        "columnMapping": mapping,
        "rows": rows,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--html", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--url", required=True)
    args = parser.parse_args()
    payload = extract_payload(Path(args.html), args.url)
    Path(args.output).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wiki规则：{len(payload['rows'])} 行")


if __name__ == "__main__":
    main()
