"""Compara los códigos de material en Excel con los del dataset y genera un informe HTML."""

from __future__ import annotations
import argparse
import html
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Sequence
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DATASET = ROOT / "dataset" / "dataset_clean"
DEFAULT_EXCEL = ROOT / "matIA-Data Set.xlsx"
DEFAULT_OUTPUT = ROOT / "outputs" / "resumen_coincidencias.html"
SPLITS = ("test", "train", "valid")
IMAGE_EXTENSIONS = {".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".webp"}

@dataclass(frozen=True)
class ImageMatch:
    split: str
    filename: str

@dataclass(frozen=True)
class MaterialRecord:
    excel_row: int
    excel_id: str
    code: str
    matches: tuple[ImageMatch, ...]

def _cell_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()

def load_material_rows(excel_path: Path) -> list[tuple[int, str, str]]:
    workbook = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        rows = worksheet.iter_rows(values_only=True)
        headers = next(rows, None)
        if headers is None:
            raise ValueError(f"El Excel no contiene filas: {excel_path}")

        column_indexes = {
            _cell_text(header).casefold(): index
            for index, header in enumerate(headers)
            if header is not None
        }
        code_index = column_indexes.get("codigo_material")
        if code_index is None:
            raise ValueError(
                f"No se encontró la columna 'codigo_material' en {excel_path}"
            )
        id_index = column_indexes.get("id_material")

        records: list[tuple[str, str]] = []
        for excel_row, row in enumerate(rows, start=2):
            code = _cell_text(row[code_index]) if code_index < len(row) else ""
            excel_id = (
                _cell_text(row[id_index])
                if id_index is not None and id_index < len(row)
                else ""
            )
            if code or excel_id:
                records.append((excel_row, excel_id, code))
        if not records:
            raise ValueError(f"No se encontraron registros en {excel_path}")
        return records
    finally:
        workbook.close()

def collect_images(dataset_dir: Path) -> dict[str, list[Path]]:
    images_by_split: dict[str, list[Path]] = {}
    for split in SPLITS:
        image_dir = dataset_dir / split / "images"
        if not image_dir.is_dir():
            raise FileNotFoundError(f"No existe la carpeta de imágenes: {image_dir}")
        images_by_split[split] = sorted(
            path
            for path in image_dir.iterdir()
            if path.is_file() and path.suffix.casefold() in IMAGE_EXTENSIONS
        )
    return images_by_split

def find_matches(
    material_rows: Sequence[tuple[int, str, str]],
    images_by_split: dict[str, list[Path]],
) -> list[MaterialRecord]:
    records: list[MaterialRecord] = []
    for excel_row, excel_id, code in material_rows:
        matches: list[ImageMatch] = []
        if code:
            code_pattern = re.compile(
                rf"(?<![A-Za-z0-9]){re.escape(code)}(?![A-Za-z0-9])",
                re.IGNORECASE,
            )
            for split, paths in images_by_split.items():
                matches.extend(
                    ImageMatch(split=split, filename=path.name)
                    for path in paths
                    if code_pattern.search(path.stem)
                )
        records.append(MaterialRecord(excel_row, excel_id, code, tuple(matches)))
    return records

def _escape(value: object) -> str:
    return html.escape(str(value), quote=True)

def _record_donut(found: int, missing: int) -> str:
    total = found + missing
    circumference = 2 * 3.141592653589793 * 48
    found_length = circumference * found / total
    return f"""
    <svg viewBox="0 0 240 180" role="img" aria-label="Registros con coincidencia: {found}; sin coincidencia: {missing}">
      <circle cx="65" cy="85" r="48" fill="none" stroke="#dc2626" stroke-width="20"/>
      <circle cx="65" cy="85" r="48" fill="none" stroke="#15803d" stroke-width="20"
        stroke-dasharray="{found_length:.2f} {circumference:.2f}" transform="rotate(-90 65 85)"/>
      <text x="65" y="82" text-anchor="middle" class="donut-total">{total}</text>
      <text x="65" y="101" text-anchor="middle" class="donut-caption">registros</text>
      <circle cx="132" cy="68" r="6" fill="#15803d"/>
      <text x="144" y="72" class="legend-label">Correctos: {found}</text>
      <circle cx="132" cy="96" r="6" fill="#dc2626"/>
      <text x="144" y="100" class="legend-label">Ausentes: {missing}</text>
    </svg>"""

def _split_chart(
    images_by_split: dict[str, list[Path]],
    matched_by_split: dict[str, int],
) -> str:
    width = 300
    max_total = max((len(paths) for paths in images_by_split.values()), default=0) or 1
    rows: list[str] = []
    for index, split in enumerate(SPLITS):
        y = 30 + index * 38
        total = len(images_by_split[split])
        matched = matched_by_split[split]
        matched_width = width * matched / max_total
        missing_width = width * (total - matched) / max_total
        rows.append(
            f'<text x="0" y="{y + 15}" class="axis-label">{_escape(split)}</text>'
            f'<rect x="65" y="{y}" width="{matched_width:.2f}" height="20" '
            f'rx="4" fill="#15803d"/>'
            f'<rect x="{65 + matched_width:.2f}" y="{y}" width="{missing_width:.2f}" '
            f'height="20" rx="4" fill="#d1d5db"/>'
            f'<text x="{73 + width * total / max_total:.2f}" y="{y + 15}" '
            f'class="axis-label">{matched} / {total}</text>'
        )
    return f"""
    <svg viewBox="0 0 440 195" role="img" aria-label="Imágenes coincidentes y no coincidentes por partición">
      {''.join(rows)}
      <rect x="65" y="150" width="12" height="12" rx="2" fill="#15803d"/>
      <text x="83" y="161" class="legend-label">Coinciden con el subconjunto</text>
      <rect x="65" y="171" width="12" height="12" rx="2" fill="#d1d5db"/>
      <text x="83" y="182" class="legend-label">Resto de imágenes</text>
    </svg>"""

def _record_table(records: Sequence[MaterialRecord], found: bool) -> str:
    selected = [record for record in records if bool(record.matches) is found]
    if not selected:
        return "<p class='empty-state'>No hay registros en esta categoría.</p>"

    rows: list[str] = []
    for record in selected:
        filenames = "".join(
            f"<li><span>{_escape(match.split)}</span> — {_escape(match.filename)}</li>"
            for match in record.matches
        )
        files_cell = (
            f"<details><summary>{len(record.matches)} imagen(es)</summary>"
            f"<ul>{filenames}</ul></details>"
            if found
            else "—"
        )
        rows.append(
            "<tr>"
            f"<td>{_escape(record.excel_id or '—')}</td>"
            f"<td>{_escape(record.code or '(vacío)')}</td>"
            f"<td>{files_cell}</td>"
            "</tr>"
        )
    return (
        "<div class='table-wrap'><table><thead><tr>"
        "<th>id_material</th><th>codigo_material</th><th>Archivos coincidentes</th>"
        "</tr></thead><tbody>"
        + "".join(rows)
        + "</tbody></table></div>"
    )

def _missing_code_panel(records: Sequence[MaterialRecord]) -> str:
    target_rows = {182, 184}
    selected = [
        record
        for record in records
        if record.excel_row in target_rows and not record.code
    ]
    rows = "".join(
        "<tr>"
        f"<td>{record.excel_row}</td>"
        f"<td>{_escape(record.excel_id or '—')}</td>"
        "<td>(vacío)</td>"
        "</tr>"
        for record in selected
    )
    table = (
        "<div class='table-wrap'><table><thead><tr>"
        "<th>Fila Excel</th><th>id_material</th><th>codigo_material</th>"
        "</tr></thead><tbody>"
        + rows
        + "</tbody></table></div>"
        if rows
        else "<p class='empty-state'>No se encontraron esas filas sin código en los datos cargados.</p>"
    )
    return f"""
  <section class="panel section data-warning">
    <h2>Filas sin codigo_material: 182 y 184</h2>
    <p>Estas filas no se pueden comparar con los nombres de las imágenes porque falta
    el código que permitiría buscarlas. Por eso cuentan como registros sin coincidencia:
    el informe no puede distinguir si falta la imagen o si falta completar el dato en el Excel.</p>
    {table}
  </section>"""

def build_html(
    records: Sequence[MaterialRecord],
    images_by_split: dict[str, list[Path]],
) -> str:
    found_records = sum(bool(record.matches) for record in records)
    missing_records = len(records) - found_records
    found_percent = 100 * found_records / len(records)
    missing_percent = 100 - found_percent

    matched_paths_by_split: dict[str, set[str]] = {split: set() for split in SPLITS}
    for record in records:
        for match in record.matches:
            matched_paths_by_split[match.split].add(match.filename)
    matched_by_split = {
        split: len(matched_paths_by_split[split]) for split in SPLITS
    }
    image_count = sum(len(paths) for paths in images_by_split.values())
    matched_image_count = sum(matched_by_split.values())
    image_percent = 100 * matched_image_count / image_count if image_count else 0

    split_rows = "".join(
        "<tr>"
        f"<td>{_escape(split)}</td>"
        f"<td>{len(images_by_split[split])}</td>"
        f"<td>{matched_by_split[split]}</td>"
        f"<td>{len(images_by_split[split]) - matched_by_split[split]}</td>"
        "</tr>"
        for split in SPLITS
    )
    generated_at = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %Z")

    return f"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Resumen de coincidencias del dataset MatIA</title>
  <style>
    :root {{ color-scheme: light; font-family: "Segoe UI", Arial, sans-serif; color: #172033; background: #f3f5f8; }}
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; padding: 32px 18px; }}
    main {{ max-width: 1120px; margin: auto; }}
    h1 {{ margin-bottom: 8px; font-size: clamp(1.7rem, 4vw, 2.4rem); }}
    h2 {{ margin: 0 0 16px; font-size: 1.2rem; }}
    p {{ color: #596579; line-height: 1.55; }}
    .subtitle {{ margin-top: 0; }}
    .cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 16px; margin: 24px 0; }}
    .charts {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; margin: 24px 0; }}
    .card, .panel {{ background: white; border: 1px solid #e3e7ed; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px #1720330a; }}
    .data-warning {{ border-left: 4px solid #d97706; }}
    .metric {{ display: block; font-size: 2rem; font-weight: 700; margin-bottom: 4px; }}
    .good {{ color: #15803d; }} .bad {{ color: #b91c1c; }}
    figure {{ margin: 0; }} svg {{ display: block; width: 100%; height: auto; overflow: visible; }}
    .donut-total {{ font-size: 24px; font-weight: 700; fill: #172033; }}
    .donut-caption, .legend-label, .axis-label {{ font-size: 12px; fill: #596579; }}
    .table-wrap {{ overflow-x: auto; }}
    table {{ width: 100%; border-collapse: collapse; text-align: left; }}
    th, td {{ border-bottom: 1px solid #e8ebf0; padding: 11px 12px; vertical-align: top; }}
    th {{ color: #596579; font-size: .88rem; }}
    details ul {{ margin: 10px 0 4px; padding-left: 20px; }}
    details li {{ margin: 5px 0; overflow-wrap: anywhere; }}
    details li span {{ color: #596579; }}
    .section {{ margin: 28px 0; }}
    .empty-state {{ padding: 8px 0; }}
    @media (max-width: 760px) {{ .charts {{ grid-template-columns: 1fr; }} }}
    @media (max-width: 560px) {{ body {{ padding: 20px 12px; }} .card, .panel {{ padding: 16px; }} }}
  </style>
</head>
<body>
<main>
  <h1>Resumen de coincidencias del dataset</h1>
  <p class="subtitle">Comparación del campo <strong>codigo_material</strong> del Excel
  con los códigos incluidos en los nombres de imágenes de test, train y valid.</p>

  <section class="cards" aria-label="Indicadores">
    <div class="card"><span class="metric">{len(records)}</span>Registros del Excel</div>
    <div class="card"><span class="metric good">{found_records} ({found_percent:.2f}%)</span>Registros correctos</div>
    <div class="card"><span class="metric bad">{missing_records} ({missing_percent:.2f}%)</span>Registros ausentes</div>
    <div class="card"><span class="metric">{matched_image_count} / {image_count}</span>
      Imágenes coincidentes ({image_percent:.2f}%)</div>
  </section>

  <section class="charts" aria-label="Gráficas">
    <article class="panel"><h2>Proporción de registros</h2>
      <figure>{_record_donut(found_records, missing_records)}</figure>
      <p>{found_records} de {len(records)} códigos aparecen en al menos un nombre de imagen.</p>
    </article>
    <article class="panel"><h2>Coincidencias por partición</h2>
      <figure>{_split_chart(images_by_split, matched_by_split)}</figure>
      <p>Las barras muestran archivos con código coincidente frente al total de imágenes.</p>
    </article>
  </section>

  {_missing_code_panel(records)}

  <section class="panel section">
    <h2>Resumen de imágenes por partición</h2>
    <div class="table-wrap"><table><thead><tr><th>Partición</th><th>Total</th>
      <th>Coincidentes</th><th>Sin coincidencia</th></tr></thead>
      <tbody>{split_rows}</tbody></table></div>
  </section>

  <section class="panel section">
    <h2>Registros correctos ({found_records})</h2>
    <p>Un registro se considera correcto cuando su codigo_material aparece como token
    independiente en uno o más nombres de archivo. Abre cada fila para ver los archivos.</p>
    {_record_table(records, True)}
  </section>

</main>
</body>
</html>
"""

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Compara los códigos del Excel con nombres de imágenes YOLO."
    )
    parser.add_argument("--excel", type=Path, default=DEFAULT_EXCEL, help="Archivo .xlsx")
    parser.add_argument(
        "--dataset",
        type=Path,
        default=DEFAULT_DATASET,
        help="Carpeta raíz del dataset que contiene test, train y valid",
    )
    parser.add_argument(
        "--output", type=Path, default=DEFAULT_OUTPUT, help="Ruta del informe HTML"
    )
    args = parser.parse_args()

    material_rows = load_material_rows(args.excel)
    images_by_split = collect_images(args.dataset)
    records = find_matches(material_rows, images_by_split)
    report = build_html(records, images_by_split)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(report, encoding="utf-8")

    found = sum(bool(record.matches) for record in records)
    print(f"Registros correctos: {found}/{len(records)} ({100 * found / len(records):.2f}%)")
    print(f"Registros ausentes: {len(records) - found}/{len(records)}")
    print(f"Informe generado: {args.output.resolve()}")

if __name__ == "__main__":
    main()