"""
Одноразовый скрипт: добавляет пункты 5.6 и 5.7 в шаблон
templates/contract_template.docx после пункта 5.5, перед пустым
параграфом-разделителем раздела 6.

Запуск (один раз):
    python scripts/add_contract_points_5_6_5_7.py

Скрипт идемпотентный: если пункты 5.6/5.7 уже присутствуют,
повторный запуск ничего не делает.
"""
from __future__ import annotations

import copy
import os
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = ROOT / "templates" / "contract_template.docx"

POINT_5_6 = (
    "5.6. В рамках лицензирования скважины, для получения санитарно-"
    "эпидемиологического заключения, необходимо выполнять сезонный "
    "мониторинг химических анализов воды из скважины согласно постановлению "
    "Правительства РФ от 29.11.2023 №2029. Сезонный мониторинг химических "
    "анализов не входит в состав работ. Полная комплектность необходимых "
    "компонентов для определения качества воды будет известна после "
    "согласования программы производственного контроля. Будет составлено "
    "дополнительное соглашение на состав этих работ. Заказчик может "
    "выполнить мониторинг самостоятельно, но при нарушении сроков или "
    "комплектности мониторинга химических анализов, срок выполнения "
    "договора может быть увеличен на срок необходимый для выполнения "
    "условий согласованного сезонного мониторинга."
)

POINT_5_7 = (
    "5.7. Согласование плана мероприятий с землепользователями, попадающими "
    "во второй и третий пояса ЗСО в соответствии с СанПиН 2.1.4.1100-02 "
    "п. 1.12. осуществляется заказчиком самостоятельно. Подрядчик "
    "предоставляет Заказчику необходимые бланки для сбора подписей."
)


def find_paragraph_starting_with(paragraphs, prefix: str):
    for p in paragraphs:
        if (p.text or "").strip().startswith(prefix):
            return p
    return None


def build_paragraph_like(reference_p, text: str):
    """
    Клонирует XML параграфа-образца (со всеми его pPr/run-properties)
    и подставляет новый текст в первый run. Все остальные run'ы удаляются.
    """
    new_p = copy.deepcopy(reference_p._element)

    runs = new_p.findall(qn("w:r"))
    if not runs:
        # На всякий случай — добавляем пустой run
        r = OxmlElement("w:r")
        new_p.append(r)
        runs = [r]

    first_run = runs[0]
    for extra in runs[1:]:
        new_p.remove(extra)

    # Удаляем старый текст/переносы из первого run, оставляем только rPr.
    for child in list(first_run):
        if child.tag in (qn("w:t"), qn("w:br"), qn("w:tab")):
            first_run.remove(child)

    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    first_run.append(t)

    return new_p


def main() -> int:
    if not TEMPLATE_PATH.exists():
        print(f"Template not found: {TEMPLATE_PATH}", file=sys.stderr)
        return 1

    doc = Document(str(TEMPLATE_PATH))
    paragraphs = doc.paragraphs

    p_5_5 = find_paragraph_starting_with(paragraphs, "5.5.")
    if p_5_5 is None:
        print("ERROR: не нашёл пункт 5.5 в шаблоне, шаблон мог измениться.",
              file=sys.stderr)
        return 2

    existing_5_6 = find_paragraph_starting_with(paragraphs, "5.6.")
    existing_5_7 = find_paragraph_starting_with(paragraphs, "5.7.")

    if existing_5_6 is not None and existing_5_7 is not None:
        print("Пункты 5.6 и 5.7 уже есть в шаблоне — ничего не делаю.")
        return 0

    # Вставляем в обратном порядке (5.7, затем 5.6) сразу после 5.5,
    # чтобы итоговый порядок получился 5.5 → 5.6 → 5.7.
    ref_elem = p_5_5._element

    if existing_5_7 is None:
        new_p_5_7 = build_paragraph_like(p_5_5, POINT_5_7)
        ref_elem.addnext(new_p_5_7)
        print("Добавлен пункт 5.7")

    if existing_5_6 is None:
        new_p_5_6 = build_paragraph_like(p_5_5, POINT_5_6)
        ref_elem.addnext(new_p_5_6)
        print("Добавлен пункт 5.6")

    doc.save(str(TEMPLATE_PATH))
    print(f"Сохранено: {TEMPLATE_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
