"""
Оценка риска клиентов по таблице Excel (НС ПОД/ФТ).

Запуск из командной строки:
    python risk_scoring.py Макет_таблицы_риска.xlsx
    python risk_scoring.py вход.xlsx -o результат.xlsx

Использование из кода сайта:
    from risk_scoring import assess_file
    results = assess_file("вход.xlsx", "результат.xlsx")

Программа:
  1. находит в таблице строку заголовков и колонки по их названиям
     (порядок колонок не важен, лишние колонки игнорируются);
  2. проверяет каждое значение (числа, «+ / −», ИНН с контрольной суммой, ОКВЭД);
  3. определяет профиль деятельности по ОКВЭД и берёт отраслевые нормы;
  4. начисляет баллы по критериям, учитывает стоп-факторы и совокупность
     признаков фиктивности;
  5. записывает балл, уровень риска и обоснование в серые колонки и
     добавляет лист «Подробный отчёт».

Все веса, пороги и нормы — в risk_config.py.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

import risk_config as cfg

# ---------------------------------------------------------------------------
# Описание колонок таблицы
# ---------------------------------------------------------------------------
# key, варианты заголовка, тип значения, обязательность
COLUMNS = [
    ("name",          ["наименование"],                                  "text",    True),
    ("inn",           ["инн"],                                           "inn",     True),
    ("okved",         ["оквэд (основной)", "оквэд"],                     "okved",   True),
    ("age",           ["возраст компании, лет", "возраст, лет", "возраст"], "age",  True),
    ("mass_address",  ["массовый адрес"],                                "flag",    True),
    ("mass_founder",  ["массовый учредитель"],                           "flag",    True),
    ("mass_director", ["массовый директор"],                             "flag",    True),
    ("unreliable",    ["недостоверные сведения в егрюл"],                "flag",    True),
    ("disqualified",  ["дисквалификация директора"],                     "flag",    True),
    ("changes",       ["смены руководства / учредителей, кол-во", "смены руководства"], "int", True),
    ("staff",         ["сотрудники, чел.", "сотрудники"],                "int",     True),
    ("revenue",       ["выручка, млн ₽", "выручка"],                     "number",  True),
    ("loss_years",    ["убыточные годы подряд"],                         "int",     True),
    ("taxes",         ["уплаченные налоги, млн ₽", "уплаченные налоги"], "number",  True),
    ("revenue_jumps", ["резкие изменения выручки"],                      "flag",    True),
    ("tax_debt",      ["налоговая задолженность"],                       "flag",    True),
    ("no_website",    ["нет сайта и публичной информации", "нет сайта"], "flag",    True),
    ("country",       ["страна регистрации"],                            "text",    True),
    ("fixed_assets",  ["основные средства, млн ₽", "основные средства"], "number",  False),
]
# Признаки, которые есть только во внутренних данных банка (несоответствие ОКВЭД
# платежам, рисковые контрагенты, транзит, наличные), в методику не входят:
# в открытых источниках их проверить нельзя.
LABELS = {
    "name": "Наименование", "inn": "ИНН", "okved": "ОКВЭД", "age": "Возраст",
    "mass_address": "Массовый адрес", "mass_founder": "Массовый учредитель",
    "mass_director": "Массовый директор", "unreliable": "Недостоверные сведения",
    "disqualified": "Дисквалификация", "changes": "Смены руководства",
    "staff": "Сотрудники", "revenue": "Выручка", "loss_years": "Убыточные годы",
    "taxes": "Уплаченные налоги", "revenue_jumps": "Резкие изменения выручки",
    "tax_debt": "Налоговая задолженность",
    "no_website": "Нет сайта", "country": "Страна", "fixed_assets": "Основные средства",
}
RESULT_COLUMNS = {"score": "итоговый балл", "level": "уровень риска", "reason": "обоснование"}

LEVEL_COLORS = {"Низкий": "C6EFCE", "Средний": "FFEB9C", "Высокий": "FFC7CE"}

TRUE_WORDS = {"+", "да", "есть", "yes", "y", "true", "1", "✓", "✔", "v"}
FALSE_WORDS = {"-", "−", "–", "—", "нет", "no", "n", "false", "0", "х", "x"}


# ---------------------------------------------------------------------------
# Структуры результата
# ---------------------------------------------------------------------------
@dataclass
class Finding:
    """Один сработавший критерий."""
    criterion: str      # номер критерия по методике, например «8»
    text: str           # человекочитаемое описание
    points: int
    stop: bool = False  # стоп-фактор: сразу «Высокий»

    @property
    def label(self) -> str:
        return f"№{self.criterion}" if self.criterion.isdigit() else self.criterion


@dataclass
class ClientResult:
    row: int
    name: str
    inn: str
    profile: str = ""
    score: int = 0
    level: str = ""
    findings: list[Finding] = field(default_factory=list)
    missing: list[str] = field(default_factory=list)   # незаполненные обязательные поля
    warnings: list[str] = field(default_factory=list)  # ошибки и замечания по вводу

    @property
    def stop_factors(self) -> list[Finding]:
        return [f for f in self.findings if f.stop]

    def reason(self) -> str:
        parts = [f"Профиль: {self.profile}."]
        if self.stop_factors:
            parts.append("СТОП-ФАКТОР: " + "; ".join(f.text for f in self.stop_factors) + ".")
        regular = [f for f in self.findings if not f.stop]
        if regular:
            parts.append("Признаки: " + "; ".join(
                f"{f.label} {f.text} (+{f.points})" for f in regular) + ".")
        elif not self.stop_factors:
            parts.append("Признаков риска не выявлено.")
        if self.missing:
            parts.append("Не заполнено: " + ", ".join(self.missing) + ".")
        if self.warnings:
            parts.append("Замечания: " + "; ".join(self.warnings) + ".")
        return " ".join(parts)


# ---------------------------------------------------------------------------
# Разбор значений ячеек
# ---------------------------------------------------------------------------
def fmt(x: float, digits: int = 1) -> str:
    """Число по-русски: 195 300,5"""
    return f"{x:,.{digits}f}".replace(",", "\u00a0").replace(".", ",")


def years(n: int) -> str:
    if n % 10 == 1 and n % 100 != 11:
        return f"{n} год"
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return f"{n} года"
    return f"{n} лет"


def _norm(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().lower().replace("ё", "е")


def _is_blank(v: Any) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def parse_flag(v: Any) -> bool | None:
    if _is_blank(v):
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    s = _norm(v)
    if s in TRUE_WORDS:
        return True
    if s in FALSE_WORDS:
        return False
    raise ValueError(f"ожидается «+» или «−», а введено «{v}»")


def parse_number(v: Any, integer: bool = False) -> float | int | None:
    if _is_blank(v):
        return None
    if isinstance(v, bool):
        raise ValueError(f"ожидается число, а введено «{v}»")
    if isinstance(v, (int, float)):
        x = float(v)
    else:
        s = str(v).lower().replace("\xa0", "").replace(" ", "")
        s = re.sub(r"(млн|руб\.?|₽|чел\.?|лет|год(а)?)$", "", s).replace(",", ".")
        try:
            x = float(s)
        except ValueError:
            raise ValueError(f"ожидается число, а введено «{v}»") from None
    if x < 0:
        raise ValueError(f"значение не может быть отрицательным ({v})")
    if integer:
        if abs(x - round(x)) > 1e-9:
            raise ValueError(f"ожидается целое число, а введено «{v}»")
        return int(round(x))
    return x


def parse_age(v: Any) -> float | None:
    """Возраст в годах. Можно ввести число лет или дату регистрации."""
    if isinstance(v, (datetime, date)):
        d = v.date() if isinstance(v, datetime) else v
        if d > date.today():
            raise ValueError("дата регистрации в будущем")
        return (date.today() - d).days / 365.25
    if isinstance(v, str) and re.fullmatch(r"\s*\d{1,2}\.\d{1,2}\.\d{4}\s*", v):
        d = datetime.strptime(v.strip(), "%d.%m.%Y").date()
        return parse_age(d)
    return parse_number(v)


def parse_inn(v: Any) -> tuple[str, str | None]:
    """Возвращает (ИНН, замечание или None). Проверяет контрольные цифры."""
    if _is_blank(v):
        return "", None
    s = str(int(v)) if isinstance(v, (int, float)) else re.sub(r"\s", "", str(v))
    if isinstance(v, (int, float)) and len(s) in (9, 11):   # Excel съел ведущий ноль
        s = s.zfill(len(s) + 1)
    if not s.isdigit() or len(s) not in (10, 12):
        return s, f"ИНН «{s}» должен состоять из 10 (юрлицо) или 12 (ИП) цифр"
    d = [int(c) for c in s]

    def check(coefs: list[int]) -> int:
        return sum(a * b for a, b in zip(coefs, d)) % 11 % 10

    if len(s) == 10:
        ok = check([2, 4, 10, 3, 5, 9, 4, 6, 8]) == d[9]
    else:
        ok = (check([7, 2, 4, 10, 3, 5, 9, 4, 6, 8]) == d[10]
              and check([3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8]) == d[11])
    return s, None if ok else f"ИНН {s} не проходит проверку контрольной суммы — вероятно, опечатка"


OKVED_RE = re.compile(r"^\d{2}(\.\d{1,2}(\.\d{1,2})?)?$")


def parse_okved(v: Any) -> tuple[list[str], list[str]]:
    """Ячейка может содержать один код или несколько через запятую/точку с запятой:
    первый — основной, остальные — дополнительные. Возвращает (коды, замечания)."""
    if _is_blank(v):
        return [], []
    notes = []
    if isinstance(v, (int, float)):
        raw = [repr(float(v)).rstrip("0").rstrip(".") if isinstance(v, float) else str(v)]
        notes.append("ОКВЭД введён числом — Excel мог отбросить конечный ноль; "
                     "лучше задать ячейке текстовый формат")
    else:
        raw = [p for p in re.split(r"[,;\s]+", str(v).strip()) if p]
    codes = []
    for p in raw:
        p = p.strip().rstrip(".")
        if re.fullmatch(r"\d\.\d+.*", p):          # «1.11» → «01.11»
            p = "0" + p
        if OKVED_RE.match(p):
            codes.append(p)
        else:
            notes.append(f"код ОКВЭД «{p}» не распознан")
    return codes, notes


def profile_for(code: str) -> str:
    """Профиль по самому длинному совпадающему префиксу кода."""
    parts = code.split(".")
    for i in range(len(parts), 0, -1):
        key = ".".join(parts[:i])
        if key in cfg.OKVED_TO_PROFILE:
            return cfg.OKVED_TO_PROFILE[key]
    return cfg.DEFAULT_PROFILE


def high_risk_code(code: str) -> str | None:
    for key, descr in cfg.HIGH_RISK_CODES.items():
        if code == key or code.startswith(key + "."):
            return descr
    return None


def okved_section(code: str) -> str | None:
    n = int(code[:2])
    for lo, hi, letter in cfg.OKVED_SECTIONS:
        if lo <= n <= hi:
            return letter
    return None


def country_status(country: str) -> str:
    """'home' | 'black' | 'grey' | 'offshore' | 'foreign'"""
    c = _norm(country).removeprefix("республика ").strip()
    if c in {cfg.HOME_COUNTRY, "рф", "russia", "российская федерация"}:
        return "home"
    if c in cfg.FATF_BLACK:
        return "black"
    if c in cfg.FATF_GREY:
        return "grey"
    if c in cfg.OFFSHORE:
        return "offshore"
    return "foreign"


# ---------------------------------------------------------------------------
# Поиск колонок в листе
# ---------------------------------------------------------------------------
def find_header_row(ws: Worksheet, max_scan: int = 20) -> int:
    for r in range(1, min(ws.max_row, max_scan) + 1):
        for c in range(1, ws.max_column + 1):
            if _norm(ws.cell(r, c).value or "") == "наименование":
                return r
    raise ValueError("Не найдена строка заголовков (ячейка «Наименование»)")


def map_columns(ws: Worksheet, header_row: int) -> tuple[dict[str, int], dict[str, int]]:
    headers = {_norm(ws.cell(header_row, c).value): c
               for c in range(1, ws.max_column + 1) if not _is_blank(ws.cell(header_row, c).value)}
    cols, missing_cols = {}, []
    for key, variants, _, mandatory in COLUMNS:
        col = next((headers[_norm(v)] for v in variants if _norm(v) in headers), None)
        if col:
            cols[key] = col
        elif mandatory:
            missing_cols.append(variants[0])
    if missing_cols:
        raise ValueError("В таблице нет обязательных колонок: " + ", ".join(missing_cols))
    res_cols = {k: headers[v] for k, v in RESULT_COLUMNS.items() if v in headers}
    return cols, res_cols


def data_rows(ws: Worksheet, header_row: int, cols: dict[str, int]) -> list[int]:
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        first = ws.cell(r, 1).value
        if isinstance(first, str) and (first.startswith("•") or _norm(first).startswith("как заполнять")):
            break                                      # дошли до легенды
        inn_cell = ws.cell(r, cols["inn"]).value
        # строки шапки «тип данных» и «подсказка» под заголовком
        if r <= header_row + 3 and isinstance(inn_cell, str) and re.search(r"[а-яa-z]", inn_cell.lower()):
            continue
        if all(_is_blank(ws.cell(r, c).value) for c in cols.values()):
            continue
        rows.append(r)
    return rows


# ---------------------------------------------------------------------------
# Оценка одного клиента
# ---------------------------------------------------------------------------
def read_client(ws: Worksheet, r: int, cols: dict[str, int]) -> tuple[dict[str, Any], ClientResult]:
    raw = {k: ws.cell(r, c).value for k, c in cols.items()}
    inn, inn_note = parse_inn(raw.get("inn"))
    res = ClientResult(row=r, name=str(raw.get("name") or "").strip(), inn=inn)
    if inn_note:
        res.warnings.append(inn_note)

    v: dict[str, Any] = {}
    for key, _, kind, mandatory in COLUMNS:
        if key not in cols:
            v[key] = None
            continue
        x = raw[key]
        try:
            if kind == "flag":
                val = parse_flag(x)
                if val is None and cfg.FLAG_BLANK_IS_NO:
                    val = False
            elif kind == "number":
                val = parse_number(x)
            elif kind == "int":
                val = parse_number(x, integer=True)
            elif kind == "age":
                val = parse_age(x)
            elif kind == "okved":
                val, notes = parse_okved(x)
                res.warnings.extend(notes)
                val = val or None
            elif kind == "inn":
                val = inn or None
            else:
                val = None if _is_blank(x) else str(x).strip()
        except ValueError as e:
            res.warnings.append(f"{LABELS[key]}: {e}")
            val = None
        v[key] = val
        if mandatory and val is None:
            res.missing.append(LABELS[key])
    return v, res


def assess(v: dict[str, Any], res: ClientResult) -> ClientResult:
    add = lambda crit, text, key, stop=False: res.findings.append(  # noqa: E731
        Finding(crit, text, cfg.W[key], stop))

    # --- Профиль деятельности ---------------------------------------------
    codes: list[str] = v["okved"] or []
    main = codes[0] if codes else None
    res.profile = profile_for(main) if main else cfg.DEFAULT_PROFILE + " (ОКВЭД не указан)"
    prof = cfg.PROFILES[profile_for(main) if main else cfg.DEFAULT_PROFILE]

    if main:
        hr = high_risk_code(main)
        if prof["high_risk"] or hr:
            add("ОКВЭД", f"рисковый основной вид деятельности ({hr or res.profile})", "high_risk_main")
        extra_hr = sorted({high_risk_code(c) or profile_for(c) for c in codes[1:]
                           if high_risk_code(c) or cfg.PROFILES[profile_for(c)]["high_risk"]})
        if extra_hr:
            add("ОКВЭД", "рисковые дополнительные виды: " + ", ".join(extra_hr), "high_risk_extra")
        sections = {okved_section(c) for c in codes} - {None}
        if len(sections) > cfg.DIVERSE_SECTIONS_LIMIT:
            add("ОКВЭД", f"несвязанные направления деятельности ({len(sections)} разделов ОКВЭД)",
                "too_diverse")

    # --- А. Регистрационные данные -----------------------------------------
    age = v["age"]
    if age is not None:
        if age < cfg.AGE_YOUNG:
            add("1", f"компания моложе 1 года ({fmt(age)} г.)", "age_lt_1")
        elif age < cfg.AGE_FAIRLY_YOUNG:
            add("1", f"компания моложе 2 лет ({fmt(age)} г.)", "age_lt_2")
    if v["mass_address"]:
        add("2", "массовый адрес регистрации", "mass_address")
    if v["mass_founder"]:
        add("3", "массовый учредитель", "mass_founder")
    if v["mass_director"]:
        add("4", "массовый директор", "mass_director")
    if v["unreliable"]:
        add("5", "недостоверные сведения в ЕГРЮЛ", "unreliable_egrul")
    if v["disqualified"]:
        add("6", "директор дисквалифицирован", "disqualified", stop=True)
    ch = v["changes"]
    if ch is not None:
        if ch > cfg.CHANGES_SEVERE:
            add("7", f"очень частая смена руководства/учредителей ({ch})", "changes_gt_5")
        elif ch > cfg.CHANGES_THRESHOLD:
            add("7", f"частая смена руководства/учредителей ({ch})", "changes_gt_2")

    # --- Б. Финансовые показатели и персонал (8, 9, 10, 11, 15) ------------
    staff, rev = v["staff"], v["revenue"]
    if staff is not None and rev is not None:
        if staff == 0 and rev > 0:
            add("8", f"нет сотрудников при выручке {fmt(rev)} млн ₽", "no_staff_revenue")
        elif staff > 0:
            rpe, norm = rev / staff, prof["rev_per_emp_max"]
            if rpe > norm * cfg.REV_PER_EMP_EXTREME_FACTOR:
                add("8", f"выручка на сотрудника {fmt(rpe)} млн ₽ — в {fmt(rpe / norm, 1)} раза выше "
                         f"нормы профиля ({fmt(norm, 0)} млн)", "rev_per_emp_extreme")
            elif rpe > norm:
                add("8", f"выручка на сотрудника {fmt(rpe)} млн ₽ выше нормы профиля ({fmt(norm, 0)} млн)",
                    "rev_per_emp_high")
    if staff is not None and not (staff == 0 and (rev or 0) > 0) and staff < prof["min_staff"]:
        add("15", f"персонала недостаточно для профиля ({staff} при минимуме {prof['min_staff']})",
            "staff_below_min")

    fa = v["fixed_assets"]
    if fa is not None and fa == 0 and prof["needs_fa"]:
        add("9", "нет основных средств, хотя для профиля они необходимы", "no_fixed_assets")

    loss = v["loss_years"]
    if loss is not None:
        if loss >= cfg.LOSS_YEARS_SEVERE:
            add("10", f"убыточна {years(loss)} подряд", "loss_3_years")
        elif loss >= cfg.LOSS_YEARS_RISK:
            add("10", f"убыточна {years(loss)} подряд", "loss_2_years")

    taxes = v["taxes"]
    if taxes is not None and rev:
        burden, avg = taxes / rev * 100, prof["avg_tax_burden"]
        if burden < avg / 2:
            add("11", f"налоговая нагрузка {fmt(burden, 2)}% — меньше половины среднеотраслевой "
                      f"({fmt(avg)}%)", "tax_below_half")
        elif burden < avg:
            add("11", f"налоговая нагрузка {fmt(burden, 2)}% ниже среднеотраслевой ({fmt(avg)}%)",
                "tax_below_avg")
    elif taxes is not None and rev == 0 and taxes == 0 and age is not None and age >= 1:
        res.warnings.append("нулевые выручка и налоги — деятельность фактически не ведётся")

    if v["revenue_jumps"]:
        add("12", "резкие изменения выручки", "revenue_jumps")
    if v["tax_debt"]:
        add("13", "налоговая задолженность", "tax_debt")

    # --- В. Реальная деятельность -----------------------------------------
    if v["no_website"]:
        add("17", "нет сайта и публичной информации", "no_website")

    # --- Г. Страновые ------------------------------------------------------
    if v["country"]:
        st = country_status(v["country"])
        if st == "black":
            res.findings.append(Finding("20", f"страна в чёрном списке FATF ({v['country']})", 0, stop=True))
        elif st == "grey":
            add("20", f"страна в сером списке FATF ({v['country']})", "country_grey")
        elif st == "offshore":
            add("20", f"офшорная юрисдикция ({v['country']})", "country_offshore")

    # --- Совокупность признаков фиктивности --------------------------------
    shell = [v["mass_address"], v["mass_founder"], v["mass_director"], v["unreliable"],
             v["tax_debt"], v["no_website"],
             staff == 0 if staff is not None else False,
             fa == 0 and prof["needs_fa"] if fa is not None else False,
             age is not None and age < cfg.AGE_YOUNG]
    n_shell = sum(bool(x) for x in shell)
    if n_shell >= cfg.SHELL_COMBO_MIN:
        add("Σ", f"совокупность признаков фиктивной компании ({n_shell})", "shell_combo")

    # --- Итог --------------------------------------------------------------
    res.score = sum(f.points for f in res.findings)
    res.level = score_to_level(res.score)
    if res.stop_factors:
        res.level = "Высокий"
    elif res.missing and _rank(res.level) < _rank(cfg.MIN_LEVEL_IF_MANDATORY_MISSING):
        res.level = cfg.MIN_LEVEL_IF_MANDATORY_MISSING
        res.warnings.append(f"уровень повышен до «{res.level}», т.к. не все обязательные поля проверены")
    return res


def _rank(level: str) -> int:
    return [name for _, name in cfg.LEVELS].index(level)


def score_to_level(score: int) -> str:
    level = cfg.LEVELS[0][1]
    for threshold, name in cfg.LEVELS:
        if score >= threshold:
            level = name
    return level


# ---------------------------------------------------------------------------
# Обработка файла целиком
# ---------------------------------------------------------------------------
def assess_file(path_in: str | Path, path_out: str | Path | None = None,
                sheet: str | None = None) -> list[ClientResult]:
    path_in = Path(path_in)
    path_out = Path(path_out) if path_out else path_in.with_name(path_in.stem + "_результат.xlsx")

    wb = load_workbook(path_in)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    header_row = find_header_row(ws)
    cols, res_cols = map_columns(ws, header_row)

    # колонки результата: если их нет — добавляем в конец
    for key, title in [("score", "Итоговый балл"), ("level", "Уровень риска"), ("reason", "Обоснование")]:
        if key not in res_cols:
            c = ws.max_column + 1
            ws.cell(header_row, c, title).font = Font(name="Arial", bold=True, size=10)
            res_cols[key] = c

    results, seen_inn = [], {}
    for r in data_rows(ws, header_row, cols):
        v, res = read_client(ws, r, cols)
        if res.inn:
            if res.inn in seen_inn:
                res.warnings.append(f"ИНН повторяется (уже есть в строке {seen_inn[res.inn]})")
            seen_inn.setdefault(res.inn, r)
        assess(v, res)
        results.append(res)

        ws.cell(r, res_cols["score"], res.score)
        lc = ws.cell(r, res_cols["level"], res.level)
        lc.fill = PatternFill("solid", fgColor=LEVEL_COLORS.get(res.level, "FFFFFF"))
        lc.font = Font(name="Arial", size=10, bold=True)
        lc.alignment = Alignment(horizontal="center", vertical="center")
        rc = ws.cell(r, res_cols["reason"], res.reason())
        rc.alignment = Alignment(wrap_text=True, vertical="top")
        rc.font = Font(name="Arial", size=9)

    write_report(wb, results)
    wb.save(path_out)
    return results


def write_report(wb, results: list[ClientResult]) -> None:
    title = "Подробный отчёт"
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title)
    bold = Font(name="Arial", bold=True, size=10)
    norm = Font(name="Arial", size=10)
    head = ["Строка", "Наименование", "ИНН", "Профиль", "Критерий", "Описание", "Баллы"]
    for i, h in enumerate(head, 1):
        c = ws.cell(1, i, h)
        c.font = bold
        c.fill = PatternFill("solid", fgColor="D9E1F2")
    r = 2
    for res in results:
        lines = [(f.criterion, ("СТОП-ФАКТОР: " if f.stop else "") + f.text, f.points) for f in res.findings]
        lines += [("—", "Не заполнено: " + m, "") for m in res.missing]
        lines += [("!", "Замечание: " + w, "") for w in res.warnings]
        if not lines:
            lines = [("—", "Признаков риска не выявлено", 0)]
        for crit, text, pts in lines:
            for i, val in enumerate([res.row, res.name, res.inn, res.profile, crit, text, pts], 1):
                ws.cell(r, i, val).font = norm
            r += 1
        c = ws.cell(r, 6, f"ИТОГО: {res.score} баллов → {res.level}")
        c.font = bold
        c.fill = PatternFill("solid", fgColor=LEVEL_COLORS.get(res.level, "FFFFFF"))
        ws.cell(r, 7, res.score).font = bold
        r += 2
    for col, w in zip("ABCDEFG", [8, 28, 13, 24, 9, 80, 8]):
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Оценка риска клиентов по таблице Excel (ПОД/ФТ)")
    p.add_argument("input", help="Excel-файл с таблицей клиентов")
    p.add_argument("-o", "--output", help="куда сохранить результат (по умолчанию <имя>_результат.xlsx)")
    p.add_argument("--sheet", help="имя листа (по умолчанию первый)")
    args = p.parse_args(argv)
    try:
        results = assess_file(args.input, args.output, args.sheet)
    except (ValueError, FileNotFoundError, KeyError) as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        return 1
    for res in results:
        print(f"[строка {res.row}] {res.name or '(без названия)'}: {res.score} б. → {res.level}")
        for f in res.findings:
            print(f"    {'СТОП ' if f.stop else ''}{f.label}: {f.text} (+{f.points})")
        for m in res.missing:
            print(f"    не заполнено: {m}")
        for w in res.warnings:
            print(f"    ! {w}")
    print(f"\nОбработано клиентов: {len(results)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
