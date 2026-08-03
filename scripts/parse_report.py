"""
parse_report.py v5.1.1

XML级解析 Word

支持:
- paragraph
- heading
- image
- chart
- table

输出:
examples/output/raw_structure.json

"""

from pathlib import Path
import argparse
import zipfile
import json
import re
import shutil

from lxml import etree



# ==========================
# 路径
# ==========================

ROOT_DIR = (
    Path(__file__)
    .resolve()
    .parent
    .parent
)


INPUT_DIR = (
    ROOT_DIR
    /
    "examples"
    /
    "raw"
)


OUTPUT_DIR = (
    ROOT_DIR
    /
    "examples"
    /
    "output"
)


OUTPUT_JSON = (
    OUTPUT_DIR
    /
    "raw_structure.json"
)


INPUT_PATH = list(
    INPUT_DIR.glob("*.docx")
)[0]



# ==========================
# namespace
# ==========================

NS = {

    "w":
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",

    "a":
    "http://schemas.openxmlformats.org/drawingml/2006/main",

    "r":
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",

    "c":
    "http://schemas.openxmlformats.org/drawingml/2006/chart"

}



# ==========================
# 标题识别
# ==========================

def detect_heading(text):

    text = text.strip()


    if re.match(
        r"^[一二三四五六七八九十]+、",
        text
    ):
        return "heading1"


    if re.match(
        r"^[（(]\s*[一二三四五六七八九十0-9]+\s*[）)]",
        text
    ):
        return "heading2"


    if re.match(
        r"^\d+[\.、]",
        text
    ):
        return "heading3"


    return None


def chinese_number(value):
    digits = "零一二三四五六七八九"
    if value < 10:
        return digits[value]
    if value == 10:
        return "十"
    if value < 20:
        return "十" + digits[value % 10]
    if value < 100:
        result = digits[value // 10] + "十"
        return result + (digits[value % 10] if value % 10 else "")
    return str(value)


def load_numbering(temp_dir):
    path = temp_dir / "word" / "numbering.xml"
    if not path.is_file():
        return {}
    root = etree.parse(str(path)).getroot()
    definitions = {}
    for num in root.xpath("./w:num", namespaces=NS):
        num_id = num.get("{%s}numId" % NS["w"])
        abstract_id = num.xpath(
            "string(w:abstractNumId/@w:val)", namespaces=NS
        )
        abstract = root.xpath(
            "./w:abstractNum[@w:abstractNumId=$abstract_id]",
            namespaces=NS,
            abstract_id=abstract_id,
        )
        if not abstract:
            continue
        levels = {}
        for level in abstract[0].xpath("./w:lvl", namespaces=NS):
            ilvl = int(level.get("{%s}ilvl" % NS["w"], "0"))
            levels[ilvl] = {
                "format": level.xpath("string(w:numFmt/@w:val)", namespaces=NS),
                "text": level.xpath("string(w:lvlText/@w:val)", namespaces=NS),
                "start": int(level.xpath("string(w:start/@w:val)", namespaces=NS) or 1),
            }
        for override in num.xpath("./w:lvlOverride", namespaces=NS):
            ilvl = int(override.get("{%s}ilvl" % NS["w"], "0"))
            start = override.xpath("string(w:startOverride/@w:val)", namespaces=NS)
            if start and ilvl in levels:
                levels[ilvl]["start"] = int(start)
        definitions[num_id] = levels
    return definitions


def paragraph_numbering(element, definitions, counters):
    num_id = element.xpath("string(w:pPr/w:numPr/w:numId/@w:val)", namespaces=NS)
    if not num_id or num_id not in definitions:
        return None
    ilvl = int(
        element.xpath("string(w:pPr/w:numPr/w:ilvl/@w:val)", namespaces=NS)
        or 0
    )
    level = definitions[num_id].get(ilvl)
    if not level:
        return None
    key = (num_id, ilvl)
    value = counters.get(key, level["start"] - 1) + 1
    counters[key] = value
    rendered = (
        chinese_number(value)
        if level["format"] in {"japaneseCounting", "chineseCounting"}
        else str(value)
    )
    marker = level["text"].replace(f"%{ilvl + 1}", rendered)
    return {"marker": marker, **level}


def paragraph_is_heading_like(element):
    bold = element.find("w:pPr/w:rPr/w:b", namespaces=NS)
    bold_value = bold.get("{%s}val" % NS["w"], "1") if bold is not None else "0"
    size = element.xpath("string(w:pPr/w:rPr/w:sz/@w:val)", namespaces=NS)
    style_id = element.xpath("string(w:pPr/w:pStyle/@w:val)", namespaces=NS)
    return style_id == "a9" or (
        bold_value not in {"0", "false", "off"} and int(size or 0) >= 24
    )



# ==========================
# 解压docx
# ==========================

def unzip_docx(input_path=INPUT_PATH):

    temp_dir = (
        ROOT_DIR
        /
        "_docx_tmp"
    )


    if temp_dir.exists():

        shutil.rmtree(
            temp_dir
        )


    temp_dir.mkdir()


    with zipfile.ZipFile(
        input_path,
        "r"
    ) as z:

        z.extractall(
            temp_dir
        )


    return temp_dir



# ==========================
# 导出图片/chart
# ==========================

def export_visual_assets(temp_dir, output_dir=OUTPUT_DIR):

    output_media = (
        output_dir
        /
        "images"
    )


    output_media.mkdir(
        exist_ok=True
    )


    # ----------------------
    # 图片
    # ----------------------

    media_dir = (
        temp_dir
        /
        "word"
        /
        "media"
    )


    if media_dir.exists():

        for file in media_dir.iterdir():

            if file.is_file():

                shutil.copy(
                    file,
                    output_media
                    /
                    file.name
                )



    # ----------------------
    # chart
    # ----------------------

    chart_dir = (
        temp_dir
        /
        "word"
        /
        "charts"
    )


    if chart_dir.exists():

        for file in chart_dir.iterdir():

            if (
                file.is_file()
                and file.name.startswith("chart")
                and file.suffix == ".xml"
        ):

                shutil.copy(
                    file,
                    output_media
                    /
                    file.name
            )



# ==========================
# rels
# ==========================

def load_relationships(temp_dir):

    rel_path = (
        temp_dir
        /
        "word"
        /
        "_rels"
        /
        "document.xml.rels"
    )


    tree = etree.parse(
        str(rel_path)
    )


    rels = {}


    for rel in tree.getroot():

        rels[
            rel.get("Id")
        ] = rel.get("Target")


    return rels



# ==========================
# 提取文字
# ==========================

def paragraph_text(element):

    texts = element.xpath(
        # Text inside drawings/text boxes is part of the visual object, not a
        # body paragraph.  Including it creates duplicated axis labels and
        # callouts between a caption and its chart.
        ".//w:t[not(ancestor::w:drawing) and not(ancestor::w:pict)]/text()",
        namespaces=NS
    )


    return "".join(texts).strip()


def paragraph_runs(element):
    """Return visible text runs with the raw report's emphasis semantics.

    The generator maps ordinary runs to Source Han Sans Light and emphasized
    runs to Medium.  Drawing text is deliberately excluded for the same reason
    as :func:`paragraph_text`.
    """

    paragraph_bold = element.find("w:pPr/w:rPr/w:b", namespaces=NS)
    paragraph_is_bold = (
        paragraph_bold is not None
        and paragraph_bold.get("{%s}val" % NS["w"], "1")
        not in {"0", "false", "off"}
    )
    runs = []
    for run in element.xpath(
        "./w:r[not(.//w:drawing) and not(.//w:pict)] | "
        "./w:hyperlink/w:r[not(.//w:drawing) and not(.//w:pict)]",
        namespaces=NS,
    ):
        text = "".join(run.xpath(".//w:t/text()", namespaces=NS))
        if not text:
            continue
        bold = run.find("w:rPr/w:b", namespaces=NS)
        if bold is None:
            emphasized = paragraph_is_bold
        else:
            emphasized = bold.get("{%s}val" % NS["w"], "1") not in {
                "0", "false", "off"
            }
        if runs and runs[-1]["emphasis"] == emphasized:
            runs[-1]["text"] += text
        else:
            runs.append({"text": text, "emphasis": emphasized})
    return runs



# ==========================
# 提取视觉对象
# ==========================

def extract_visual(
    element,
    rels
):

    result = []


    # 图片

    blips = element.xpath(
        ".//a:blip",
        namespaces=NS
    )


    for b in blips:

        rid = b.get(
            "{%s}embed" % NS["r"]
        )


        if rid in rels:

            result.append(
                {
                    "type":"figure",
                    "object_type":"image",
                    "path":
                    "images/"
                    +
                    Path(
                        rels[rid]
                    ).name
                }
            )



    # chart

    charts = element.xpath(
        ".//c:chart",
        namespaces=NS
    )


    for c in charts:

        rid = c.get(
            "{%s}id" % NS["r"]
        )


        if rid in rels:

            result.append(
                {
                    "type":"figure",
                    "object_type":"chart",
                    "path":
                    "images/"
                    +
                    Path(
                        rels[rid]
                    ).name
                }
            )


    return result



# ==========================
# 表格
# ==========================

def parse_table(tbl):

    rows = []


    for tr in tbl.xpath(
        ".//w:tr",
        namespaces=NS
    ):

        row = []


        for tc in tr.xpath(
            "./w:tc",
            namespaces=NS
        ):

            row.append(
                "".join(
                    tc.xpath(
                        ".//w:t/text()",
                        namespaces=NS
                    )
                )
            )


        rows.append(
            row
        )


    return rows



# ==========================
# 主解析
# ==========================

def parse(temp_dir):

    rels = load_relationships(
        temp_dir
    )


    doc_xml = (
        temp_dir
        /
        "word"
        /
        "document.xml"
    )


    tree = etree.parse(
        str(doc_xml)
    )


    body = tree.xpath(
        "//w:body",
        namespaces=NS
    )[0]


    content = []
    numbering = load_numbering(temp_dir)
    numbering_counters = {}


    for child in body:

        tag = etree.QName(
            child
        ).localname



        # paragraph

        if tag == "p":


            text = paragraph_text(
                child
            )


            if text:


                heading = detect_heading(
                    text
                )

                implicit_number = paragraph_numbering(
                    child, numbering, numbering_counters
                )
                if not heading and implicit_number and paragraph_is_heading_like(child):
                    number_format = implicit_number["format"]
                    if number_format in {"japaneseCounting", "chineseCounting"}:
                        heading = "heading2"
                    elif number_format == "decimal":
                        heading = "heading3"
                    if heading:
                        text = implicit_number["marker"] + text


                if heading:

                    content.append(
                        {
                            "type":heading,
                            "text":text,
                            "runs": paragraph_runs(child),
                        }
                    )


                elif text.startswith("图"):

                    content.append(
                        {
                            "type":
                            "figure_caption",

                            "text":
                            text
                        }
                    )


                elif text.startswith("表"):

                    content.append(
                        {
                            "type":
                            "table_caption",

                            "text":
                            text
                        }
                    )


                else:

                    content.append(
                        {
                            "type":
                            "paragraph",

                            "text":
                            text,

                            "runs":
                            paragraph_runs(child),
                        }
                    )



            content.extend(
                extract_visual(
                    child,
                    rels
                )
            )



        # table

        elif tag == "tbl":


            content.append(
                {
                    "type":"table",

                    "data":
                    parse_table(child)
                }
            )


    return content



# ==========================
# main
# ==========================

def parse_args():

    parser = argparse.ArgumentParser(
        description="Parse a DOCX into ordered report blocks"
    )

    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        default=INPUT_PATH
    )

    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_JSON
    )

    return parser.parse_args()


def main():

    args = parse_args()

    input_path = args.input

    output_json = args.output

    output_dir = output_json.parent

    print(
        "Parsing:",
        input_path.name
    )


    output_dir.mkdir(
        parents=True,
        exist_ok=True
    )


    temp_dir = unzip_docx(
        input_path
    )


    export_visual_assets(
        temp_dir,
        output_dir
    )


    result = {

        "filename":
        input_path.name,


        "content":
        parse(
            temp_dir
        )

    }


    with open(
        output_json,
        "w",
        encoding="utf-8"
    ) as f:


        json.dump(
            result,
            f,
            ensure_ascii=False,
            indent=4
        )


    print(
        "Done:"
    )


    print(
        output_json
    )



if __name__ == "__main__":

    main()
