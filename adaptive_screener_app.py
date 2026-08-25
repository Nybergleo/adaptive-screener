"""BQuant app for an adaptive Bloomberg equity screen."""

from __future__ import annotations

from dataclasses import dataclass, field
import difflib
import json
from pathlib import Path
import builtins
from datetime import datetime
import html
import re
import time
import warnings

import numpy as np
import pandas as pd

try:
    import bql
except ModuleNotFoundError:  # pragma: no cover - local/demo mode
    bql = None

try:
    import ipywidgets as widgets
    from IPython.display import HTML, clear_output, display
except ModuleNotFoundError as exc:  # pragma: no cover - BQuant/Jupyter dependency
    raise ModuleNotFoundError(
        "AdaptiveScreenerApp requires ipywidgets and IPython. BQL is optional in demo mode."
    ) from exc


APP_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = APP_DIR / "data_cache" / "margin_stability_screener_v02"
OUTPUT_PATH = OUTPUT_DIR / "margin_stability_screen_results.csv"
PRESETS_PATH = OUTPUT_DIR / "margin_stability_presets.json"

WESTERN_EUROPE_COUNTRIES = {
    "Austria": "AT",
    "Belgium": "BE",
    "Denmark": "DK",
    "Finland": "FI",
    "France": "FR",
    "Germany": "DE",
    "Ireland": "IE",
    "Italy": "IT",
    "Luxembourg": "LU",
    "Netherlands": "NL",
    "Norway": "NO",
    "Portugal": "PT",
    "Spain": "ES",
    "Sweden": "SE",
    "Switzerland": "CH",
    "United Kingdom": "GB",
}

NORTH_AMERICA_COUNTRIES = {
    "Canada": "CA",
    "United States": "US",
}

SUPPORTED_COUNTRIES = {
    **WESTERN_EUROPE_COUNTRIES,
    **NORTH_AMERICA_COUNTRIES,
}

COUNTRY_GROUPS = {
    "western_europe": WESTERN_EUROPE_COUNTRIES,
    "north_america": NORTH_AMERICA_COUNTRIES,
}

DISPLAY_COLUMNS = [
    "id",
    "name",
    "cntry_of_domicile",
    "crncy",
    "sector",
    "industry_group",
    "ebit_margin_ltm",
    "ebita_margin_ltm",
    "ebitda_margin_ltm",
    "margin_stability_pp",
    "fwd_pe_blended_12m",
    "ebit_ltm",
    "market_cap",
    "turnover",
]

PERCENT_COLUMNS = {"ebit_margin_ltm", "ebita_margin_ltm", "ebitda_margin_ltm", "margin_stability_pp"}


@dataclass(frozen=True)
class FieldSpec:
    key: str
    label: str
    alias: str
    expression: str
    field_type: str
    definition: str
    default_min: str = ""
    default_max: str = ""
    default_text: str = ""
    default_input: bool = True
    default_display: bool = True
    format_type: str = "number"
    aliases: tuple[str, ...] = field(default_factory=tuple)

    def bql_expression(self, *, currency: str, n_quarters: int) -> str:
        fpo_start = -(int(n_quarters) - 1)
        return self.expression.format(currency=currency, fpo_start=fpo_start, n_quarters=n_quarters)


IDENTITY_FIELDS = [
    FieldSpec("name", "Company name", "name", "name()", "text", "Bloomberg company/security name.", format_type="text"),
    FieldSpec("cntry_of_domicile", "Country of domicile", "cntry_of_domicile", "cntry_of_domicile", "text", "Bloomberg country-of-domicile country code.", format_type="text"),
    FieldSpec("crncy", "Trading currency", "crncy", "crncy", "text", "Bloomberg security currency.", format_type="text"),
    FieldSpec("sector", "BICS sector", "sector", "classification_name(BICS, 1)", "text", "Bloomberg BICS level 1 sector.", format_type="text", aliases=("classification_name_bics_1",)),
    FieldSpec("industry_group", "BICS industry group", "industry_group", "classification_name(BICS, 2)", "text", "Bloomberg BICS level 2 industry group.", format_type="text", aliases=("classification_name_bics_2",)),
]

FIELD_CATALOG = [
    FieldSpec(
        "ebit_margin_ltm",
        "LTM EBIT / operating margin",
        "ebit_margin_ltm",
        "oper_margin(fpt=LTM)",
        "numeric",
        "Reported LTM operating margin. Bloomberg oper_margin maps to EBIT divided by revenue times 100.",
        default_min="10",
        format_type="percent_points",
        aliases=("oper_margin_fpt_ltm", "oper_margin"),
    ),
    FieldSpec(
        "ebita_margin_ltm",
        "LTM EBITA margin",
        "ebita_margin_ltm",
        "ebita_margin(fpt=LTM)",
        "numeric",
        "Reported LTM EBITA margin. Coverage can be uneven across companies.",
        default_min="11",
        format_type="percent_points",
        aliases=("ebita_margin_fpt_ltm", "ebita_margin"),
    ),
    FieldSpec(
        "ebitda_margin_ltm",
        "LTM EBITDA margin",
        "ebitda_margin_ltm",
        "ebitda_to_revenue(fpt=LTM)",
        "numeric",
        "Reported LTM EBITDA divided by revenue.",
        format_type="percent_points",
        aliases=("ebitda_margin_fpt_ltm", "ebitda_margin", "ebitda_to_revenue_fpt_ltm", "ebitda_to_revenue"),
    ),
    FieldSpec(
        "margin_stability_pp",
        "Quarterly EBIT margin stability",
        "margin_stability_pp",
        "std(oper_margin(fpt=Q, fpo=range({fpo_start},0)))",
        "numeric",
        "Standard deviation of quarterly operating margin over the selected trailing quarter window. Lower means more stable margins.",
        default_max="4",
        format_type="percent_points",
        aliases=("std_oper_margin_fpt_q_fpo_range",),
    ),
    FieldSpec(
        "fwd_pe_blended_12m",
        "Blended 12M forward P/E",
        "fwd_pe_blended_12m",
        "pe_ratio(fpt=BT, fpo=1)",
        "numeric",
        "Bloomberg blended 12-month forward consensus P/E.",
        default_max="25",
        aliases=("pe_ratio_fpt_bt_fpo_1",),
    ),
    FieldSpec(
        "ebit_ltm",
        "LTM EBIT",
        "ebit_ltm",
        "ebit(fpt=LTM, currency={currency})",
        "numeric",
        "Reported LTM EBIT in the selected currency.",
        aliases=("ebit_fpt_ltm_currency",),
    ),
    FieldSpec(
        "market_cap",
        "Market capitalization",
        "market_cap",
        "cur_mkt_cap(currency={currency})",
        "numeric",
        "Current market capitalization in the selected currency.",
        default_min="500000000",
        aliases=("cur_mkt_cap_currency",),
    ),
    FieldSpec(
        "turnover",
        "Turnover",
        "turnover",
        "turnover(currency={currency})",
        "numeric",
        "Bloomberg turnover in the selected currency.",
        aliases=("turnover_currency",),
    ),
    FieldSpec(
        "revenue_ltm",
        "LTM revenue",
        "revenue_ltm",
        "sales_rev_turn(fpt=LTM, currency={currency})",
        "numeric",
        "Reported LTM revenue in the selected currency.",
        aliases=("sales_rev_turn_fpt_ltm_currency", "sales_rev_turn"),
    ),
    FieldSpec(
        "free_cash_flow_yield",
        "Free cash flow yield",
        "free_cash_flow_yield",
        "free_cash_flow_yield",
        "numeric",
        "Bloomberg free cash flow yield.",
        format_type="percent_points",
    ),
    FieldSpec(
        "dividend_yield",
        "Dividend yield",
        "dividend_yield",
        "dividend_yield",
        "numeric",
        "Bloomberg indicated dividend yield.",
        format_type="percent_points",
    ),
    FieldSpec(
        "return_on_capital",
        "Return on capital",
        "return_on_capital",
        "return_on_capital",
        "numeric",
        "Bloomberg return on capital metric.",
        format_type="percent_points",
    ),
    FieldSpec(
        "sector_filter",
        "BICS sector",
        "sector",
        "classification_name(BICS, 1)",
        "text",
        "Bloomberg BICS level 1 sector. Use exact names or comma-separated values.",
        default_display=False,
        format_type="text",
        aliases=("classification_name_bics_1",),
    ),
    FieldSpec(
        "industry_group_filter",
        "BICS industry group",
        "industry_group",
        "classification_name(BICS, 2)",
        "text",
        "Bloomberg BICS level 2 industry group. Use exact names or comma-separated values.",
        default_display=False,
        format_type="text",
        aliases=("classification_name_bics_2",),
    ),
]

DEFAULT_FIELD_KEYS = [
    "ebit_margin_ltm",
    "ebita_margin_ltm",
    "ebitda_margin_ltm",
    "margin_stability_pp",
    "fwd_pe_blended_12m",
    "ebit_ltm",
    "market_cap",
    "turnover",
]

DEFAULT_STABILITY_QUARTERS = 12

PARAMETER_DEFINITIONS = {
    "universe_mode": (
        "Geographic scope applied with Bloomberg cntry_of_domicile country codes."
    ),
    "countries": "Geographic restriction using Bloomberg cntry_of_domicile country codes.",
    "currency": "Currency override for absolute value fields such as EBIT, market cap, and turnover. Ratio fields are unchanged.",
    "n_quarters": (
        "Field-level operating-margin stability window. The app computes std(oper_margin(fpt=Q, fpo=range(-N+1,0))) "
        "server-side in BQL for selected fields that expose a Window control."
    ),
    "min_market_cap": "Minimum current market capitalization in the selected currency: cur_mkt_cap(currency=<currency>).",
    "max_market_cap": "Maximum current market capitalization in the selected currency: cur_mkt_cap(currency=<currency>).",
    "min_turnover": "Minimum Bloomberg turnover in the selected currency: turnover(currency=<currency>).",
    "max_turnover": "Maximum Bloomberg turnover in the selected currency: turnover(currency=<currency>).",
    "min_ebit": "Minimum absolute EBIT in the selected currency. The app uses ebit(fpt=LTM, currency=<currency>).",
    "max_ebit": "Maximum absolute EBIT in the selected currency. The app uses ebit(fpt=LTM, currency=<currency>).",
    "min_ebit_margin": (
        "Minimum EBIT / operating margin. Bloomberg oper_margin maps to EBIT divided by revenue times 100. "
        "The app uses oper_margin(fpt=LTM)."
    ),
    "max_ebit_margin": (
        "Maximum EBIT / operating margin. Bloomberg oper_margin maps to EBIT divided by revenue times 100. "
        "The app uses oper_margin(fpt=LTM)."
    ),
    "min_ebita_margin": (
        "Minimum EBITA margin. Bloomberg ebita_margin is EBIT plus amortisation of intangibles, divided by "
        "revenue times 100. Coverage can be uneven."
    ),
    "max_ebita_margin": (
        "Maximum EBITA margin. Bloomberg ebita_margin is EBIT plus amortisation of intangibles, divided by "
        "revenue times 100. Coverage can be uneven."
    ),
    "min_margin_std": (
        "Minimum operating-margin volatility in percentage points. Computed as the standard deviation of "
        "quarterly oper_margin over the selected trailing window."
    ),
    "max_margin_std": (
        "Maximum operating-margin volatility in percentage points. Lower values mean more stable margins. "
        "Computed inline with BQL std()."
    ),
    "min_pe": (
        "Minimum forward blended P/E. The app uses pe_ratio(fpt=BT, fpo=1), a rolling 12-month forward "
        "consensus P/E."
    ),
    "max_pe": (
        "Maximum forward blended P/E. The app uses pe_ratio(fpt=BT, fpo=1), a rolling 12-month forward "
        "consensus P/E."
    ),
    "search": "Local table search after Bloomberg has returned the screen results.",
    "sort_by": "Local table sort column after Bloomberg has returned the screen results.",
    "ascending": "Sort direction for the local result table.",
    "top_n": "Number of rows displayed locally. This does not change the Bloomberg query.",
    "show_logs": "Show the generated BQL query and execution details for debugging.",
}


def log(verbose: bool, *args, **kwargs) -> None:
    if verbose:
        builtins.print(*args, **kwargs)


def status(label: str, detail: str = ""):
    timestamp = datetime.now().strftime("%H:%M:%S")
    suffix = f" | {detail}" if detail else ""
    builtins.print(f"[{timestamp}] {label}{suffix}", flush=True)


class NotebookProgress:
    def __init__(self, title: str):
        self.title = title
        self.title_label = widgets.Label(value=title)
        self.progress = widgets.IntProgress(value=0, min=0, max=100, description="")
        self.detail_label = widgets.Label(value="")
        self.box = widgets.VBox([self.title_label, self.progress, self.detail_label])
        self.displayed = False

    def update(self, step: int, total: int, label: str, detail: str = "", running: bool = True):
        pct = 0 if total == 0 else max(0, min(100, step / total * 100))
        state = "Running" if running else "Done"
        self.progress.value = int(pct)
        self.title_label.value = f"{self.title}: {state} {step}/{total} - {label}"
        self.detail_label.value = detail
        if not self.displayed:
            display(self.box)
            self.displayed = True


def parse_universe_text(text: str) -> list[str]:
    tickers = []
    for chunk in text.replace(",", "\n").splitlines():
        ticker = chunk.strip()
        if ticker and not ticker.startswith("#"):
            tickers.append(ticker)
    return list(dict.fromkeys(tickers))


def ids_to_bql(ids: list[str]) -> str:
    return "[" + ",".join(f"'{ticker}'" for ticker in ids) + "]"


def parse_optional_number(value) -> float | None:
    text = str(value).strip().replace(",", "").replace("_", "")
    if text == "":
        return None
    return float(text)


def add_threshold(conditions: list[str], expression: str, operator: str, value) -> None:
    parsed = parse_optional_number(value)
    if parsed is not None:
        conditions.append(f"{expression} {operator} {parsed}")


def set_widget_tooltip(widget, text: str) -> None:
    """Attach hover help across common ipywidgets/BQuant versions."""
    for attr in ("description_tooltip", "tooltip"):
        try:
            setattr(widget, attr, text)
        except Exception:
            pass


def help_button(*definition_keys: str):
    text = "\n\n".join(PARAMETER_DEFINITIONS[key] for key in definition_keys)
    return help_button_text(text)


def help_button_text(text: str):
    detail = widgets.HTML(
        (
            "<div style='font-size:12px;color:#444;line-height:1.35;margin-top:6px;"
            "white-space:normal;overflow:visible;max-width:100%;'>"
            f"{html.escape(text)}</div>"
        ),
        layout=widgets.Layout(display="none", width="auto"),
    )
    button = widgets.Button(
        description="?",
        tooltip=text,
        layout=widgets.Layout(width="32px"),
    )

    def toggle_description(_button=None):
        detail.layout.display = "" if detail.layout.display == "none" else "none"

    button.on_click(toggle_description)
    return widgets.VBox([button, detail], layout=widgets.Layout(width="auto", overflow="visible"))


def with_help(widget, *definition_keys: str):
    return widgets.HBox([widget, help_button(*definition_keys)])


def bql_quote(value: str) -> str:
    return "'" + value.replace("'", "\\'") + "'"


def parse_text_values(value: str) -> list[str]:
    values = []
    for chunk in value.replace("\n", ",").split(","):
        item = chunk.strip()
        if item:
            values.append(item)
    return list(dict.fromkeys(values))


def safe_alias(value: str, fallback: str = "custom_field") -> str:
    alias = re.sub(r"[^a-zA-Z0-9_]+", "_", value.strip().lower()).strip("_")
    return alias or fallback


def field_search_score(spec: FieldSpec, query: str) -> float:
    haystack = " ".join([spec.label, spec.alias, spec.expression, spec.definition]).lower()
    query = query.lower().strip()
    if not query:
        return 1.0
    if query in haystack:
        return 2.0 + len(query) / max(len(haystack), 1)
    return difflib.SequenceMatcher(None, query, haystack).ratio()


def format_field_option(spec: FieldSpec) -> str:
    return f"{spec.label} ({spec.alias}) - {spec.expression}"


def catalog_by_key() -> dict[str, FieldSpec]:
    return {spec.key: spec for spec in FIELD_CATALOG}


def field_uses_window(spec: FieldSpec) -> bool:
    return "{fpo_start}" in spec.expression or "{n_quarters}" in spec.expression


def field_spec_to_payload(spec: FieldSpec) -> dict:
    return {
        "key": spec.key,
        "label": spec.label,
        "alias": spec.alias,
        "expression": spec.expression,
        "field_type": spec.field_type,
        "definition": spec.definition,
        "default_min": spec.default_min,
        "default_max": spec.default_max,
        "default_text": spec.default_text,
        "default_input": spec.default_input,
        "default_display": spec.default_display,
        "format_type": spec.format_type,
        "aliases": list(spec.aliases),
    }


def field_spec_from_payload(payload: dict) -> FieldSpec:
    return FieldSpec(
        key=str(payload.get("key", "")),
        label=str(payload.get("label", "")),
        alias=safe_alias(str(payload.get("alias", payload.get("label", "custom_field")))),
        expression=str(payload.get("expression", "")),
        field_type=str(payload.get("field_type", "numeric")),
        definition=str(payload.get("definition", "")),
        default_min=str(payload.get("default_min", "")),
        default_max=str(payload.get("default_max", "")),
        default_text=str(payload.get("default_text", "")),
        default_input=bool(payload.get("default_input", True)),
        default_display=bool(payload.get("default_display", True)),
        format_type=str(payload.get("format_type", "number")),
        aliases=tuple(str(alias) for alias in payload.get("aliases", [])),
    )


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    x = df.copy()
    x.columns = (
        x.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace(r"[^a-z0-9_]+", "_", regex=True)
        .str.strip("_")
    )
    return x


def first_valid(series: pd.Series):
    values = series.dropna()
    return values.iloc[0] if not values.empty else np.nan


def collapse_company_level(df: pd.DataFrame) -> pd.DataFrame:
    x = df.copy()
    if "id" not in x.columns:
        return x
    value_cols = [col for col in x.columns if col != "id"]
    return x.groupby("id", as_index=False)[value_cols].agg(first_valid)


def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols = list(df.columns)
    for candidate in candidates:
        if candidate in cols:
            return candidate
    for candidate in candidates:
        for col in cols:
            if col.startswith(candidate + "_"):
                return col
    for candidate in candidates:
        for col in cols:
            if candidate in col:
                return col
    return None


def add_aliases(df: pd.DataFrame) -> pd.DataFrame:
    x = df.copy()
    aliases = {
        "name": ["name"],
        "sector": ["classification_name_bics_1"],
        "industry_group": ["classification_name_bics_2"],
        "ebit_margin_ltm": ["oper_margin_fpt_ltm", "oper_margin"],
        "ebita_margin_ltm": ["ebita_margin_fpt_ltm", "ebita_margin"],
        "ebitda_margin_ltm": ["ebitda_margin_fpt_ltm", "ebitda_margin", "ebitda_to_revenue_fpt_ltm", "ebitda_to_revenue"],
        "margin_stability_pp": ["std_oper_margin_fpt_q_fpo_range"],
        "fwd_pe_blended_12m": ["pe_ratio_fpt_bt_fpo_1"],
        "ebit_ltm": ["ebit_fpt_ltm_currency"],
        "market_cap": ["cur_mkt_cap_currency"],
        "turnover": ["turnover_currency"],
    }
    for target, candidates in aliases.items():
        source = find_col(x, [target] + candidates)
        if source is not None and target not in x.columns:
            x[target] = x[source]
    return x


def run_bql(query: str) -> pd.DataFrame:
    if bql is None:
        raise RuntimeError("BQL is not available. Enable Demo mode to review the UI without Bloomberg.")
    bq = bql.Service()
    response = bq.execute(query)
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="In a future version, the Index constructor will not infer numeric dtypes",
            category=FutureWarning,
            module="bql",
        )
        warnings.filterwarnings(
            "ignore",
            message="The 'combined_df' function is deprecated",
            category=PendingDeprecationWarning,
        )
        return bql.combined_df(response).reset_index()


def format_table(df: pd.DataFrame, percent_columns: set[str] | None = None):
    percent_columns = percent_columns or PERCENT_COLUMNS
    formatters = {}
    for col in df.columns:
        if col in percent_columns:
            formatters[col] = "{:.1f}"
        elif pd.api.types.is_float_dtype(df[col]):
            formatters[col] = "{:,.2f}"
    try:
        return df.style.format(formatters, na_rep="").set_table_styles(
            [
                {"selector": "th", "props": [("text-align", "left"), ("font-weight", "600")]},
                {"selector": "td", "props": [("text-align", "left")]},
            ]
        )
    except AttributeError:
        return df


class MarginStabilityScreenerApp:
    """Run and review a reusable adaptive BQL equity screen."""

    def __init__(self, demo_mode: bool = False):
        self.df = pd.DataFrame()
        self.filtered = pd.DataFrame()
        self.session_presets: dict[str, dict] = {}
        self.progress = NotebookProgress("Adaptive screener")

        self.universe_mode = widgets.Dropdown(
            options=[
                ("Western Europe", "western_europe"),
                ("North America", "north_america"),
            ],
            value="western_europe",
            description="Geography",
            layout=widgets.Layout(width="260px"),
        )
        self.index_ticker = widgets.Text(
            value="SXXP Index",
            description="Index",
            layout=widgets.Layout(width="260px"),
        )
        self.custom_tickers = widgets.Textarea(
            value="",
            placeholder="One Bloomberg ticker per line, or comma-separated",
            description="Tickers",
            rows=4,
            layout=widgets.Layout(width="520px"),
        )
        self.custom_tickers_dropdown = widgets.Accordion(children=[self.custom_tickers])
        self.custom_tickers_dropdown.set_title(0, "Custom tickers")
        self.custom_tickers_dropdown.selected_index = None
        self.countries = widgets.SelectMultiple(
            options=[(name, code) for name, code in WESTERN_EUROPE_COUNTRIES.items()],
            value=tuple(COUNTRY_GROUPS["western_europe"].values()),
            description="Countries",
            rows=8,
            layout=widgets.Layout(width="280px"),
        )
        self.currency = widgets.Dropdown(
            options=["EUR", "GBP", "USD", "SEK", "DKK", "NOK", "CHF"],
            value="EUR",
            description="Currency",
            layout=widgets.Layout(width="210px"),
        )
        self.n_quarters = widgets.Dropdown(
            options=[("8 quarters", 8), ("12 quarters", 12), ("20 quarters", 20)],
            value=DEFAULT_STABILITY_QUARTERS,
            description="Window",
            layout=widgets.Layout(width="220px"),
        )
        self.catalog = catalog_by_key()
        self.selected_fields: dict[str, FieldSpec] = {}
        self.field_widgets: dict[str, dict[str, widgets.Widget]] = {}
        self.custom_counter = 0
        self.field_search = widgets.Text(
            value="",
            placeholder="Search labels",
            description="Field",
            layout=widgets.Layout(width="100%", max_width="520px"),
        )
        self.field_results = widgets.Select(
            options=[],
            rows=5,
            description="Matches",
            layout=widgets.Layout(width="100%", max_width="680px", display="none"),
        )
        self.field_results_panel = widgets.Accordion(children=[self.field_results])
        self.field_results_panel.set_title(0, "Top matches")
        self.field_results_panel.selected_index = None
        self.add_field_button = widgets.Button(description="Add Field", button_style="info", layout=widgets.Layout(width="120px"))
        self.custom_label = widgets.Text(value="", placeholder="Custom label", description="Label", layout=widgets.Layout(width="260px"))
        self.custom_expression = widgets.Text(value="", placeholder="BQL expression, e.g. px_to_book_ratio", description="BQL", layout=widgets.Layout(width="420px"))
        self.custom_alias = widgets.Text(value="", placeholder="Optional output alias", description="Alias", layout=widgets.Layout(width="260px"))
        self.custom_type = widgets.Dropdown(options=[("Numeric", "numeric"), ("Text", "text")], value="numeric", description="Type", layout=widgets.Layout(width="180px"))
        self.custom_note = widgets.Textarea(value="", placeholder="Hover definition for this custom field", description="Note", rows=2, layout=widgets.Layout(width="100%", max_width="680px"))
        self.add_custom_button = widgets.Button(description="Add Custom", layout=widgets.Layout(width="120px"))
        self.custom_field_panel = widgets.Accordion(
            children=[
                widgets.VBox(
                    [
                        widgets.HBox([self.custom_label, self.custom_alias, self.custom_type]),
                        self.custom_expression,
                        self.custom_note,
                        self.add_custom_button,
                    ],
                    layout=widgets.Layout(width="100%"),
                )
            ]
        )
        self.custom_field_panel.set_title(0, "Custom Bloomberg/BQL field")
        self.custom_field_panel.selected_index = None
        self.selected_fields_box = widgets.VBox(layout=widgets.Layout(width="100%"))
        self.search = widgets.Text(value="", placeholder="Ticker or company", description="Search", layout=widgets.Layout(width="320px"))
        self.sort_by = widgets.Dropdown(options=["margin_stability_pp"], value="margin_stability_pp", description="Sort", layout=widgets.Layout(width="270px"))
        self.ascending = widgets.Checkbox(value=True, description="Ascending")
        self.top_n = widgets.IntSlider(value=50, min=10, max=300, step=10, description="Rows", continuous_update=False)
        self._syncing_page = False
        self.page_last_button = widgets.Button(description="Last", layout=widgets.Layout(width="72px", display="none"))
        self.page_number = widgets.BoundedIntText(value=1, min=1, max=1, description="Page", layout=widgets.Layout(width="150px"))
        self.page_next_button = widgets.Button(description="Next", layout=widgets.Layout(width="72px", display="none"))
        self.page_status = widgets.Label(value="")
        self.show_logs = widgets.Checkbox(value=False, description="Show logs")
        self.demo_mode = widgets.Checkbox(value=demo_mode or bql is None, description="Demo mode")
        self.demo_mode.disabled = bql is None
        if bql is None:
            set_widget_tooltip(self.demo_mode, "BQL is not available in this environment, so demo mode is forced.")
        self.run_button = widgets.Button(description="Run Screen", button_style="primary")
        self.export_button = widgets.Button(description="Export View")
        self.copy_button = widgets.Button(description="Copy View")
        self.preset_name = widgets.Text(value="", placeholder="Preset name", description="Preset", layout=widgets.Layout(width="260px"))
        self.preset_picker = widgets.Dropdown(options=[], description="Saved", layout=widgets.Layout(width="320px"))
        self.save_preset_button = widgets.Button(description="Save Preset", button_style="success", layout=widgets.Layout(width="120px"))
        self.load_preset_button = widgets.Button(description="Load Preset", layout=widgets.Layout(width="120px"))
        self.delete_preset_button = widgets.Button(description="Delete Preset", layout=widgets.Layout(width="120px"))
        self.preset_status = widgets.Label(value="")
        self.export_presets_button = widgets.Button(description="Export JSON", layout=widgets.Layout(width="110px"))
        self.import_presets_button = widgets.Button(description="Import JSON", layout=widgets.Layout(width="110px"))
        self.preset_json = widgets.Textarea(
            value="",
            placeholder="Preset JSON appears here when exported. Paste preset JSON here before importing.",
            rows=5,
            layout=widgets.Layout(width="100%", max_width="760px"),
        )
        self.preset_json_panel = widgets.Accordion(children=[self.preset_json])
        self.preset_json_panel.set_title(0, "Preset JSON import/export")
        self.preset_json_panel.selected_index = None
        self.status_label = widgets.Label(value="Status: Idle. Set filters and click Run Screen.")
        self.query_output = widgets.Output()
        self.copy_output = widgets.Output()
        self.table_output = widgets.Output()

        self._apply_parameter_tooltips()
        self.run_button.on_click(self._run)
        self.export_button.on_click(self._export)
        self.copy_button.on_click(self._copy)
        self.save_preset_button.on_click(self._save_preset)
        self.load_preset_button.on_click(self._load_selected_preset)
        self.delete_preset_button.on_click(self._delete_selected_preset)
        self.export_presets_button.on_click(self._export_presets_json)
        self.import_presets_button.on_click(self._import_presets_json)
        self.add_field_button.on_click(self._add_selected_catalog_field)
        self.add_custom_button.on_click(self._add_custom_field)
        self.page_last_button.on_click(self._previous_page)
        self.page_next_button.on_click(self._next_page)
        self.page_number.observe(self._page_changed, names="value")
        self.field_search.observe(self._sync_field_results, names="value")
        self.universe_mode.observe(self._sync_controls, names="value")
        for control in [self.search, self.sort_by, self.ascending, self.top_n]:
            control.observe(self._display_control_changed, names="value")
        for key in DEFAULT_FIELD_KEYS:
            self._add_field(self.catalog[key], render=False)
        self._sync_field_results()
        self._render_selected_fields()
        self._sync_preset_picker()
        self._sync_controls()

    def show(self):
        responsive_row = widgets.Layout(display="flex", flex_flow="row wrap", align_items="flex-start", gap="8px", width="100%")
        universe_panel = widgets.VBox(
            [
                widgets.HBox([self.universe_mode, self.currency], layout=responsive_row),
            ],
            layout=widgets.Layout(width="100%"),
        )
        preset_panel = widgets.VBox(
            [
                widgets.HBox([self.preset_name, self.save_preset_button], layout=responsive_row),
                widgets.HBox([self.preset_picker, self.load_preset_button, self.delete_preset_button], layout=responsive_row),
                widgets.HBox([self.export_presets_button, self.import_presets_button], layout=responsive_row),
                self.preset_json_panel,
                widgets.HTML(
                    "<span style='color:#555;font-size:12px;'>"
                    "Presets are saved as local JSON in this BQuant project. "
                    "In published BQAP apps, runtime file writes may not persist after close; "
                    "include shared presets in the project before publishing."
                    "</span>"
                ),
                self.preset_status,
            ],
            layout=widgets.Layout(width="100%"),
        )
        fields_panel = widgets.VBox(
            [
                widgets.HTML("<b>Presets</b>"),
                preset_panel,
                widgets.HTML("<b>Universe</b>"),
                universe_panel,
                widgets.HBox([self.field_search, self.add_field_button], layout=responsive_row),
                self.field_results_panel,
                self.custom_field_panel,
                widgets.HTML("<b>Selected filters and output fields</b>"),
                self.selected_fields_box,
            ],
            layout=widgets.Layout(width="100%"),
        )
        display_panel = widgets.VBox(
            [
                widgets.HBox([self.search, self.sort_by, self.ascending, self.top_n, self.show_logs, self.demo_mode], layout=responsive_row),
                widgets.HBox([self.page_last_button, self.page_number, self.page_next_button, self.page_status], layout=responsive_row),
                widgets.HBox([self.run_button, self.export_button, self.copy_button], layout=responsive_row),
                self.status_label,
            ],
            layout=widgets.Layout(width="100%"),
        )
        fields_section = widgets.Accordion(children=[fields_panel])
        fields_section.set_title(0, "Fields")
        fields_section.selected_index = 0
        display_section = widgets.Accordion(children=[display_panel])
        display_section.set_title(0, "Display and run")
        display_section.selected_index = 0
        controls = widgets.VBox(
            [
                fields_section,
                display_section,
            ]
        )
        display(controls, self.query_output, self.copy_output, self.table_output)

    def _sync_field_results(self, _change=None):
        query = self.field_search.value
        ranked = sorted(FIELD_CATALOG, key=lambda spec: field_search_score(spec, query), reverse=True)
        matches = [spec for spec in ranked if not query.strip() or field_search_score(spec, query) >= 0.18]
        self.field_results.options = [(format_field_option(spec), spec.key) for spec in matches[:5]]
        if self.field_results.options and self.field_results.value is None:
            self.field_results.value = self.field_results.options[0][1]
        has_query = bool(query.strip())
        self.field_results.layout.display = "" if has_query else "none"
        self.field_results_panel.layout.display = "" if has_query else "none"
        self.field_results_panel.selected_index = 0 if has_query else None

    def _add_selected_catalog_field(self, _button=None):
        key = self.field_results.value
        if key in self.catalog:
            self._add_field(self.catalog[key])

    def _add_custom_field(self, _button=None):
        expression = self.custom_expression.value.strip()
        label = self.custom_label.value.strip() or expression
        note = self.custom_note.value.strip()
        if not expression:
            self._set_status("Field missing", "Enter a BQL expression before adding a custom field.", "error")
            return
        if not note:
            self._set_status("Definition missing", "Add a hover note for the custom field.", "error")
            return

        self.custom_counter += 1
        alias = safe_alias(self.custom_alias.value or label, fallback=f"custom_field_{self.custom_counter}")
        key = f"custom_{self.custom_counter}_{alias}"
        spec = FieldSpec(
            key=key,
            label=label,
            alias=alias,
            expression=expression,
            field_type=self.custom_type.value,
            definition=note,
            format_type="text" if self.custom_type.value == "text" else "number",
        )
        self._add_field(spec)
        self.custom_label.value = ""
        self.custom_expression.value = ""
        self.custom_alias.value = ""
        self.custom_note.value = ""

    def _add_field(self, spec: FieldSpec, render: bool = True):
        if spec.key in self.selected_fields:
            return
        self.selected_fields[spec.key] = spec
        widgets_for_field: dict[str, widgets.Widget] = {
            "input": widgets.Checkbox(value=spec.default_input, description="Input", indent=False, layout=widgets.Layout(width="80px")),
            "display": widgets.Checkbox(value=spec.default_display, description="Output", indent=False, layout=widgets.Layout(width="90px")),
        }
        if field_uses_window(spec):
            widgets_for_field["n_quarters"] = widgets.Dropdown(
                options=[("8Q", 8), ("12Q", 12), ("20Q", 20)],
                value=DEFAULT_STABILITY_QUARTERS,
                description="Window",
                layout=widgets.Layout(width="150px"),
            )
        if spec.field_type == "numeric":
            widgets_for_field["min"] = widgets.Text(value=spec.default_min, placeholder="Min", description="Min", layout=widgets.Layout(width="160px"))
            widgets_for_field["max"] = widgets.Text(value=spec.default_max, placeholder="Max", description="Max", layout=widgets.Layout(width="160px"))
        else:
            widgets_for_field["text"] = widgets.Text(value=spec.default_text, placeholder="Exact value(s), comma-separated", description="Equals", layout=widgets.Layout(width="320px"))
        self.field_widgets[spec.key] = widgets_for_field
        if render:
            self._render_selected_fields()

    def _remove_field(self, key: str):
        self.selected_fields.pop(key, None)
        self.field_widgets.pop(key, None)
        self._render_selected_fields()

    def _render_selected_fields(self):
        rows = []
        row_layout = widgets.Layout(display="flex", flex_flow="row wrap", align_items="center", gap="8px", width="100%", overflow="visible")
        card_layout = widgets.Layout(border="1px solid #d2d5d9", padding="8px", margin="0 0 6px 0", width="100%", overflow="visible")
        for key, spec in self.selected_fields.items():
            remove_button = widgets.Button(description="Remove", layout=widgets.Layout(width="86px"))
            remove_button.on_click(lambda _button, field_key=key: self._remove_field(field_key))
            title = widgets.HTML(f"<b>{html.escape(spec.label)}</b><br><span style='color:#555;font-size:12px;'>{html.escape(spec.expression)}</span>", layout=widgets.Layout(width="280px"))
            help_box = help_button_text(spec.definition)
            help_button_widget = help_box.children[0]
            help_detail_widget = help_box.children[1]
            controls = [title, help_button_widget, self.field_widgets[key]["input"], self.field_widgets[key]["display"]]
            if "n_quarters" in self.field_widgets[key]:
                controls.append(self.field_widgets[key]["n_quarters"])
            if spec.field_type == "numeric":
                controls.extend([self.field_widgets[key]["min"], self.field_widgets[key]["max"]])
            else:
                controls.append(self.field_widgets[key]["text"])
            controls.append(remove_button)
            rows.append(widgets.VBox([widgets.HBox(controls, layout=row_layout), help_detail_widget], layout=widgets.Layout(width="100%", overflow="visible")))
        if not rows:
            rows = [widgets.HTML("<span style='color:#555;'>No selected fields. Add a catalog or custom field before running.</span>")]
        self.selected_fields_box.children = tuple(widgets.Box([row], layout=card_layout) for row in rows)

    def _read_presets(self) -> dict:
        presets = dict(self.session_presets)
        if not PRESETS_PATH.exists():
            return presets
        try:
            with PRESETS_PATH.open("r", encoding="utf-8") as handle:
                data = json.load(handle)
        except (json.JSONDecodeError, OSError) as exc:
            self.preset_status.value = f"Could not read presets: {exc}"
            return presets
        if isinstance(data, dict):
            presets.update(data.get("presets", {}))
        return presets

    def _write_presets(self, presets: dict) -> None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        payload = {"version": 1, "presets": presets}
        tmp_path = PRESETS_PATH.with_suffix(PRESETS_PATH.suffix + ".tmp")
        with tmp_path.open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, sort_keys=True)
            handle.write("\n")
        tmp_path.replace(PRESETS_PATH)

    def _preset_store_payload(self) -> dict:
        return {"version": 1, "presets": self._read_presets()}

    def _sync_preset_picker(self) -> None:
        presets = self._read_presets()
        names = sorted(presets)
        self.preset_picker.options = [(name, name) for name in names] or [("No saved presets", "")]
        if self.preset_picker.value not in names:
            self.preset_picker.value = names[0] if names else ""

    def _current_preset_payload(self, name: str) -> dict:
        fields = []
        for key, spec in self.selected_fields.items():
            controls = self.field_widgets.get(key, {})
            item = {
                "key": key,
                "source": "catalog" if key in self.catalog else "custom",
                "spec": field_spec_to_payload(spec) if key not in self.catalog else None,
                "input": bool(controls.get("input", widgets.Checkbox(value=False)).value),
                "display": bool(controls.get("display", widgets.Checkbox(value=False)).value),
            }
            if spec.field_type == "numeric":
                item["min"] = str(controls.get("min", widgets.Text(value="")).value)
                item["max"] = str(controls.get("max", widgets.Text(value="")).value)
            else:
                item["text"] = str(controls.get("text", widgets.Text(value="")).value)
            if "n_quarters" in controls:
                item["n_quarters"] = controls["n_quarters"].value
            fields.append(item)

        return {
            "name": name,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "universe": {
                "geography": self.universe_mode.value,
                "currency": self.currency.value,
            },
            "display": {
                "search": self.search.value,
                "sort_by": self.sort_by.value,
                "ascending": self.ascending.value,
                "top_n": self.top_n.value,
                "page": self.page_number.value,
                "show_logs": self.show_logs.value,
            },
            "fields": fields,
        }

    def _save_preset(self, _button=None):
        name = self.preset_name.value.strip()
        if not name:
            self.preset_status.value = "Enter a preset name before saving."
            return
        try:
            presets = self._read_presets()
            payload = self._current_preset_payload(name)
            presets[name] = payload
            self.session_presets[name] = payload
        except Exception as exc:
            self.preset_status.value = f"Could not prepare preset: {type(exc).__name__}: {exc}"
            return
        try:
            self._write_presets(presets)
        except Exception as exc:
            self._sync_preset_picker()
            self.preset_picker.value = name
            self.preset_json.value = json.dumps({"version": 1, "presets": presets}, indent=2, sort_keys=True)
            self.preset_json_panel.selected_index = 0
            self.preset_status.value = f"Saved for this session only. Could not write JSON: {type(exc).__name__}: {exc}"
            return
        self._sync_preset_picker()
        self.preset_picker.value = name
        self.preset_status.value = f"Saved preset: {name} ({PRESETS_PATH})"

    def _load_selected_preset(self, _button=None):
        name = self.preset_picker.value
        if not name:
            self.preset_status.value = "No preset selected."
            return
        presets = self._read_presets()
        payload = presets.get(name)
        if not payload:
            self.preset_status.value = f"Preset not found: {name}"
            self._sync_preset_picker()
            return
        try:
            self._apply_preset_payload(payload)
        except Exception as exc:
            self.preset_status.value = f"Could not load preset: {type(exc).__name__}: {exc}"
            return
        self.preset_name.value = name
        self.preset_status.value = f"Loaded preset: {name}"

    def _delete_selected_preset(self, _button=None):
        name = self.preset_picker.value
        if not name:
            self.preset_status.value = "No preset selected."
            return
        presets = self._read_presets()
        if name in presets:
            del presets[name]
            self.session_presets.pop(name, None)
            try:
                self._write_presets(presets)
            except Exception as exc:
                self._sync_preset_picker()
                self.preset_json.value = json.dumps({"version": 1, "presets": presets}, indent=2, sort_keys=True)
                self.preset_json_panel.selected_index = 0
                self.preset_status.value = f"Deleted for this session only. Could not update JSON: {type(exc).__name__}: {exc}"
                return
        self._sync_preset_picker()
        self.preset_status.value = f"Deleted preset: {name}"

    def _export_presets_json(self, _button=None):
        try:
            self.preset_json.value = json.dumps(self._preset_store_payload(), indent=2, sort_keys=True)
        except Exception as exc:
            self.preset_status.value = f"Could not export presets: {type(exc).__name__}: {exc}"
            return
        self.preset_json_panel.selected_index = 0
        self.preset_status.value = "Exported presets to JSON text."

    def _import_presets_json(self, _button=None):
        raw = self.preset_json.value.strip()
        if not raw:
            self.preset_status.value = "Paste preset JSON before importing."
            return
        try:
            data = json.loads(raw)
            incoming = data.get("presets", data) if isinstance(data, dict) else {}
            if not isinstance(incoming, dict):
                raise ValueError("Preset JSON must contain an object of presets.")
            presets = self._read_presets()
            presets.update(incoming)
            self.session_presets.update(incoming)
        except Exception as exc:
            self.preset_status.value = f"Could not import presets: {type(exc).__name__}: {exc}"
            return
        try:
            self._write_presets(presets)
            self.preset_status.value = f"Imported {len(incoming)} preset(s) and saved JSON."
        except Exception as exc:
            self.preset_status.value = f"Imported {len(incoming)} preset(s) for this session only. Could not write JSON: {type(exc).__name__}: {exc}"
        self._sync_preset_picker()

    def _apply_preset_payload(self, payload: dict) -> None:
        universe = payload.get("universe", {})
        display_settings = payload.get("display", {})

        legacy_mode = universe.get("mode")
        geography = universe.get("geography")
        if geography is None:
            geography = legacy_mode if legacy_mode in COUNTRY_GROUPS else self._infer_geography(universe.get("countries", []))

        if geography in [value for _label, value in self.universe_mode.options]:
            self.universe_mode.value = geography
        self._sync_country_options()
        self.index_ticker.value = universe.get("index_ticker", self.index_ticker.value)
        self.custom_tickers.value = universe.get("custom_tickers", self.custom_tickers.value)
        valid_countries = {code for _name, code in self.countries.options}
        countries = tuple(code for code in universe.get("countries", self.countries.value) if code in valid_countries)
        if countries:
            self.countries.value = countries
        if universe.get("currency") in self.currency.options:
            self.currency.value = universe["currency"]
        preset_window = universe.get("n_quarters", DEFAULT_STABILITY_QUARTERS)

        self.selected_fields.clear()
        self.field_widgets.clear()
        for item in payload.get("fields", []):
            key = item.get("key", "")
            spec = self.catalog.get(key)
            if spec is None and item.get("spec"):
                spec = field_spec_from_payload(item["spec"])
                self.custom_counter += 1
            if spec is None or not spec.key or not spec.expression:
                continue
            self._add_field(spec, render=False)
            controls = self.field_widgets.get(spec.key, {})
            if "input" in controls:
                controls["input"].value = bool(item.get("input", spec.default_input))
            if "display" in controls:
                controls["display"].value = bool(item.get("display", spec.default_display))
            if "n_quarters" in controls:
                quarter_values = [value for _label, value in controls["n_quarters"].options]
                n_quarters = item.get("n_quarters", preset_window)
                if n_quarters in quarter_values:
                    controls["n_quarters"].value = n_quarters
            if spec.field_type == "numeric":
                if "min" in controls:
                    controls["min"].value = str(item.get("min", spec.default_min))
                if "max" in controls:
                    controls["max"].value = str(item.get("max", spec.default_max))
            elif "text" in controls:
                controls["text"].value = str(item.get("text", spec.default_text))

        self.search.value = display_settings.get("search", self.search.value)
        self.ascending.value = bool(display_settings.get("ascending", self.ascending.value))
        self.show_logs.value = bool(display_settings.get("show_logs", self.show_logs.value))
        top_n = display_settings.get("top_n", self.top_n.value)
        try:
            self.top_n.value = min(max(int(top_n), self.top_n.min), self.top_n.max)
        except (TypeError, ValueError):
            pass
        page = display_settings.get("page", self.page_number.value)
        try:
            self.page_number.value = max(1, int(page))
        except (TypeError, ValueError):
            self.page_number.value = 1
        sort_by = display_settings.get("sort_by")
        if sort_by in self.sort_by.options:
            self.sort_by.value = sort_by

        self._sync_controls()
        self._render_selected_fields()
        if not self.df.empty:
            self._sync_sort_options()
            sort_by = display_settings.get("sort_by")
            if sort_by in self.sort_by.options:
                self.sort_by.value = sort_by
            self._refresh()

    def _apply_parameter_tooltips(self) -> None:
        for attr, definition in PARAMETER_DEFINITIONS.items():
            widget = getattr(self, attr, None)
            if widget is not None:
                set_widget_tooltip(widget, definition)

    def _set_status(self, title: str, detail: str = "", tone: str = "running") -> None:
        prefix = {
            "idle": "Status",
            "running": "Running",
            "success": "Completed",
            "error": "Failed",
        }.get(tone, title)
        message = detail or title
        self.status_label.value = f"{prefix}: {message}"

    def _infer_geography(self, countries: list[str] | tuple[str, ...]) -> str:
        selected = set(countries)
        if selected and selected.issubset(set(NORTH_AMERICA_COUNTRIES.values())):
            return "north_america"
        if selected and selected.issubset(set(WESTERN_EUROPE_COUNTRIES.values())):
            return "western_europe"
        return "western_europe"

    def _country_group_for_universe(self) -> dict[str, str]:
        return COUNTRY_GROUPS.get(self.universe_mode.value, COUNTRY_GROUPS["western_europe"])

    def _sync_country_options(self) -> None:
        countries = self._country_group_for_universe()
        options = [(name, code) for name, code in countries.items()]
        current = tuple(code for code in self.countries.value if code in countries.values())
        self.countries.options = options
        self.countries.value = current or tuple(countries.values())

    def _sync_controls(self, _change=None):
        self._sync_country_options()
        self.index_ticker.layout.display = "none"
        self.custom_tickers_dropdown.layout.display = "none"
        self.custom_tickers_dropdown.selected_index = None

    def _base_universe_expression(self) -> str:
        return "equitiesUniv([ACTIVE, PRIMARY])"

    def _all_output_specs(self) -> list[FieldSpec]:
        selected = [
            spec
            for key, spec in self.selected_fields.items()
            if bool(self.field_widgets.get(key, {}).get("display", widgets.Checkbox(value=False)).value)
        ]
        selected_aliases = {spec.alias for spec in selected}
        identity = [spec for spec in IDENTITY_FIELDS if spec.alias not in selected_aliases]
        return identity + selected

    def _display_field_specs(self) -> list[FieldSpec]:
        selected = [
            spec
            for key, spec in self.selected_fields.items()
            if bool(self.field_widgets.get(key, {}).get("display", widgets.Checkbox(value=False)).value)
        ]
        selected_aliases = {spec.alias for spec in selected}
        identity = [spec for spec in IDENTITY_FIELDS if spec.alias not in selected_aliases]
        return identity + selected

    def _percent_columns(self) -> set[str]:
        cols = set(PERCENT_COLUMNS)
        for spec in self.selected_fields.values():
            if spec.format_type == "percent_points":
                cols.add(spec.alias)
        return cols

    def _field_expression(self, spec: FieldSpec) -> str:
        controls = self.field_widgets.get(spec.key, {})
        n_quarters_widget = controls.get("n_quarters")
        n_quarters = n_quarters_widget.value if n_quarters_widget is not None else DEFAULT_STABILITY_QUARTERS
        return spec.bql_expression(currency=self.currency.value, n_quarters=int(n_quarters))

    def _dynamic_aliases(self, df: pd.DataFrame) -> pd.DataFrame:
        x = df.copy()
        for spec in self._all_output_specs():
            source = find_col(x, [spec.alias, spec.key, *spec.aliases])
            if source is not None and spec.alias not in x.columns:
                x[spec.alias] = x[source]
        return x

    def _field_conditions(self, spec: FieldSpec, expression: str) -> list[str]:
        field_controls = self.field_widgets.get(spec.key, {})
        if spec.field_type == "numeric":
            conditions = []
            min_widget = field_controls.get("min")
            max_widget = field_controls.get("max")
            if min_widget is not None:
                add_threshold(conditions, expression, ">=", min_widget.value)
            if max_widget is not None:
                add_threshold(conditions, expression, "<=", max_widget.value)
            return conditions

        text_widget = field_controls.get("text")
        if text_widget is None:
            return []
        values = parse_text_values(text_widget.value)
        if not values:
            return []
        if len(values) == 1:
            return [f"{expression} == {bql_quote(values[0])}"]
        return [f"{expression} in [" + ",".join(bql_quote(value) for value in values) + "]"]

    def _demo_numeric_values(self, spec: FieldSpec, n_rows: int, rng) -> np.ndarray:
        alias = spec.alias
        if "stability" in alias or "std" in alias:
            return rng.gamma(2.0, 1.0, n_rows).clip(0.2, 7)
        if "margin" in alias or spec.format_type == "percent_points":
            return rng.normal(14, 5, n_rows).clip(-5, 35)
        if "pe" in alias:
            return rng.normal(18, 6, n_rows).clip(4, 45)
        if "market_cap" in alias:
            return rng.lognormal(mean=21.0, sigma=1.0, size=n_rows)
        if "turnover" in alias:
            return rng.lognormal(mean=18.0, sigma=1.1, size=n_rows)
        if "ebit" in alias:
            return rng.lognormal(mean=18.8, sigma=1.0, size=n_rows)
        if "revenue" in alias:
            return rng.lognormal(mean=20.0, sigma=1.0, size=n_rows)
        return rng.normal(10, 3, n_rows)

    def _demo_text_values(self, spec: FieldSpec, n_rows: int, rng) -> list[str]:
        if spec.alias == "sector":
            choices = ["Technology", "Health Care", "Industrials", "Consumer Discretionary", "Materials"]
        elif spec.alias == "industry_group":
            choices = ["Software", "Capital Goods", "Pharmaceuticals", "Retailing", "Commercial Services"]
        elif spec.alias == "crncy":
            choices = ["EUR", "GBP", "SEK", "DKK", "NOK", "CHF"]
        elif spec.alias == "cntry_of_domicile":
            choices = list(self.countries.value) or list(self._country_group_for_universe().values())
        else:
            choices = ["Sample A", "Sample B", "Sample C"]
        return [choices[int(i)] for i in rng.integers(0, len(choices), n_rows)]

    def _demo_result(self) -> pd.DataFrame:
        rng = np.random.default_rng(42)
        n_rows = 120
        countries = list(self.countries.value) or list(self._country_group_for_universe().values())
        df = pd.DataFrame(
            {
                "id": [f"DEMO{i:03d} Equity" for i in range(1, n_rows + 1)],
                "name": [f"Demo Company {i:03d}" for i in range(1, n_rows + 1)],
                "cntry_of_domicile": [countries[int(i)] for i in rng.integers(0, len(countries), n_rows)],
                "crncy": [self.currency.value for _ in range(n_rows)],
                "sector": self._demo_text_values(FieldSpec("sector", "BICS sector", "sector", "", "text", ""), n_rows, rng),
                "industry_group": self._demo_text_values(FieldSpec("industry_group", "BICS industry group", "industry_group", "", "text", ""), n_rows, rng),
            }
        )
        for spec in self.selected_fields.values():
            if spec.alias in df.columns:
                continue
            if spec.field_type == "numeric":
                df[spec.alias] = self._demo_numeric_values(spec, n_rows, rng)
            else:
                df[spec.alias] = self._demo_text_values(spec, n_rows, rng)

        for key, spec in self.selected_fields.items():
            controls = self.field_widgets.get(key, {})
            if not bool(controls.get("input", widgets.Checkbox(value=False)).value) or spec.alias not in df.columns:
                continue
            if spec.field_type == "numeric":
                min_value = parse_optional_number(controls.get("min", widgets.Text(value="")).value)
                max_value = parse_optional_number(controls.get("max", widgets.Text(value="")).value)
                if min_value is not None:
                    df = df[df[spec.alias] >= min_value]
                if max_value is not None:
                    df = df[df[spec.alias] <= max_value]
            else:
                values = parse_text_values(controls.get("text", widgets.Text(value="")).value)
                if values:
                    df = df[df[spec.alias].astype(str).isin(values)]
        return df.reset_index(drop=True)

    def _build_query(self) -> str:
        countries = list(self.countries.value)
        if not countries:
            raise ValueError("Select at least one country.")

        country_expr = "[" + ",".join(f"'{code}'" for code in countries) + "]"
        base_universe = self._base_universe_expression()

        conditions = [
            f"cntry_of_domicile in {country_expr}",
            "security_typ == 'Common Stock'",
        ]
        get_lines = []
        seen_expressions = set()
        for spec in self._all_output_specs():
            expression = self._field_expression(spec)
            if expression not in seen_expressions:
                get_lines.append(f"    {expression}")
                seen_expressions.add(expression)
        for key, spec in self.selected_fields.items():
            if bool(self.field_widgets.get(key, {}).get("input", widgets.Checkbox(value=False)).value):
                expression = self._field_expression(spec)
                conditions.extend(self._field_conditions(spec, expression))

        condition_text = "\n        and ".join(conditions)
        get_text = ",\n".join(get_lines)

        return f"""
get(
{get_text}
)
for(
    filter(
        {base_universe},
        {condition_text}
    )
)
"""

    def _run(self, _button=None):
        started_at = time.perf_counter()
        with self.query_output:
            clear_output(wait=True)
        with self.table_output:
            clear_output(wait=True)

        try:
            self._set_status("Running", "1/4 building BQL query...", "running")
            self.progress.update(1, 4, "Build BQL query", "Preparing adaptive equity screen.")
            status("START 1/4: Build BQL query")
            query = self._build_query()
            log(self.show_logs.value, query)
            status("DONE 1/4: Build BQL query")

            if self.demo_mode.value:
                self._set_status("Running", "2/4 generating demo data. Bloomberg is not queried.", "running")
                self.progress.update(2, 4, "Demo data", "Generating deterministic sample rows from the selected fields.")
                status("START 2/4: Demo data", "Skipping Bloomberg/BQL execution.")
                self.df = self._prepare_result(self._demo_result())
                status("DONE 2/4: Demo data", str(self.df.shape))
            else:
                self._set_status("Running", "2/4 querying Bloomberg. Filters are applied server-side.", "running")
                self.progress.update(2, 4, "Run BQL", "Querying Bloomberg.")
                status("START 2/4: Run BQL", "Bloomberg applies the universe and metric filters server-side here.")
                raw = run_bql(query)
                log(self.show_logs.value, raw.shape)
                status("DONE 2/4: Run BQL", str(raw.shape))

                self._set_status("Running", "3/4 normalizing Bloomberg output...", "running")
                self.progress.update(3, 4, "Normalize output", "Collapsing sparse BQL rows to one row per company.")
                status("START 3/4: Normalize output")
                df = self._dynamic_aliases(add_aliases(collapse_company_level(normalize_columns(raw))))
                self.df = self._prepare_result(df)
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            self.df.to_csv(OUTPUT_PATH, index=False)
            status("DONE 3/4: Prepare output", str(self.df.shape))

            self._set_status("Running", "4/4 rendering table and saving CSV...", "running")
            self._sync_sort_options()
            self.progress.update(4, 4, "Complete", f"Saved {OUTPUT_PATH}", running=False)
            status("DONE 4/4: Complete", f"Saved {OUTPUT_PATH}")
            self._refresh()
            elapsed = time.perf_counter() - started_at
            self._set_status("Completed", f"Runtime: {elapsed:.1f}s. Passing companies: {len(self.df):,}.", "success")
        except Exception as exc:
            elapsed = time.perf_counter() - started_at
            self.progress.update(4, 4, "Failed", f"{type(exc).__name__}: {exc}", running=False)
            self._set_status("Failed", f"Runtime: {elapsed:.1f}s. {type(exc).__name__}: {exc}", "error")
            with self.table_output:
                clear_output(wait=True)
                print(f"{type(exc).__name__}: {exc}")

    def _prepare_result(self, df: pd.DataFrame) -> pd.DataFrame:
        x = df.copy()
        ordered = [spec.alias for spec in self._all_output_specs() if spec.alias in x.columns]
        rest = [col for col in x.columns if col not in ordered]
        x = x[ordered + rest]
        if "margin_stability_pp" in x.columns:
            x = x.sort_values("margin_stability_pp", ascending=True, na_position="last")
        return x

    def _sync_sort_options(self):
        options = [spec.alias for spec in self._display_field_specs() if spec.alias in self.df.columns]
        if not options:
            options = list(self.df.columns)
        current = self.sort_by.value if self.sort_by.value in options else options[0]
        self.sort_by.options = options
        self.sort_by.value = current
        if len(self.df):
            self.top_n.max = max(10, len(self.df))
            self.top_n.value = min(self.top_n.value, self.top_n.max)

    def _apply_filters(self) -> pd.DataFrame:
        df = self.df.copy()
        if self.search.value.strip() and not df.empty:
            query = self.search.value.strip().lower()
            mask = pd.Series(False, index=df.index)
            for col in ["id", "name"]:
                if col in df.columns:
                    mask |= df[col].astype(str).str.lower().str.contains(query, na=False)
            df = df[mask]
        if self.sort_by.value in df.columns:
            df = df.sort_values(self.sort_by.value, ascending=self.ascending.value, na_position="last")
        return df

    def _display_control_changed(self, _change=None):
        self._set_page(1, refresh=False)
        self._refresh()

    def _page_size(self) -> int:
        return max(1, int(self.top_n.value))

    def _total_pages(self) -> int:
        if self.filtered.empty:
            return 1
        return max(1, int(np.ceil(len(self.filtered) / self._page_size())))

    def _set_page(self, page: int, refresh: bool = True):
        total_pages = self._total_pages()
        page = max(1, min(int(page), total_pages))
        self._syncing_page = True
        self.page_number.max = total_pages
        self.page_number.value = page
        self._syncing_page = False
        self._sync_page_controls()
        if refresh:
            self._refresh()

    def _page_changed(self, change):
        if self._syncing_page:
            return
        self._set_page(change["new"])

    def _previous_page(self, _button=None):
        self._set_page(int(self.page_number.value) - 1)

    def _next_page(self, _button=None):
        self._set_page(int(self.page_number.value) + 1)

    def _page_bounds(self) -> tuple[int, int]:
        if self.filtered.empty:
            return (0, 0)
        start = (int(self.page_number.value) - 1) * self._page_size()
        end = min(start + self._page_size(), len(self.filtered))
        return (start, end)

    def _sync_page_controls(self):
        total_pages = self._total_pages()
        current = max(1, min(int(self.page_number.value), total_pages))
        self._syncing_page = True
        self.page_number.max = total_pages
        self.page_number.value = current
        self._syncing_page = False
        self.page_last_button.layout.display = "" if current > 1 else "none"
        self.page_next_button.layout.display = "" if current < total_pages else "none"
        start, end = self._page_bounds()
        if self.filtered.empty:
            self.page_status.value = "No rows"
        else:
            self.page_status.value = f"Rows {start + 1:,}-{end:,} of {len(self.filtered):,} | {total_pages:,} pages"

    def _refresh(self, _change=None):
        if self.df.empty:
            return
        self.filtered = self._apply_filters()
        self._sync_page_controls()
        shown = self._current_display_table()
        with self.copy_output:
            clear_output(wait=True)
        with self.table_output:
            clear_output(wait=True)
            display(format_table(shown, self._percent_columns()))

    def _current_display_table(self) -> pd.DataFrame:
        start, end = self._page_bounds()
        shown = self.filtered.iloc[start:end]
        display_cols = [spec.alias for spec in self._display_field_specs() if spec.alias in shown.columns]
        if not display_cols:
            display_cols = list(shown.columns)
        return shown[display_cols]

    def _export(self, _button=None):
        if self.filtered.empty:
            with self.table_output:
                print("No filtered rows to export.")
            return
        out = OUTPUT_DIR / "margin_stability_filtered_view.csv"
        self._current_display_table().to_csv(out, index=False)
        with self.table_output:
            print(f"Exported filtered view: {out}")

    def _copy(self, _button=None):
        if self.filtered.empty:
            with self.copy_output:
                clear_output(wait=True)
                print("No filtered rows to copy.")
            return

        visible_table = self._current_display_table()
        tsv = visible_table.to_csv(sep="\t", index=False, lineterminator="\n")
        escaped_tsv = tsv.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
        textarea_value = html.escape(tsv)
        row_count = len(visible_table)
        html_block = f"""
        <div style="margin:6px 0 10px 0;">
          <button id="copy-margin-stability-view" type="button"
                  style="height:30px;border:1px solid #1d4f91;background:#1d4f91;color:#fff;border-radius:4px;padding:0 10px;cursor:pointer;">
            Copy TSV
          </button>
          <span id="copy-margin-stability-status" style="margin-left:8px;color:#555;font-size:12px;">
            {row_count:,} visible rows ready for Excel paste
          </span>
          <textarea id="copy-margin-stability-text"
                    style="width:100%;height:110px;margin-top:6px;font-family:Consolas,monospace;font-size:11px;">{textarea_value}</textarea>
        </div>
        <script>
        (() => {{
          const text = `{escaped_tsv}`;
          const button = document.getElementById("copy-margin-stability-view");
          const status = document.getElementById("copy-margin-stability-status");
          const textarea = document.getElementById("copy-margin-stability-text");
          textarea.focus();
          textarea.select();
          button.addEventListener("click", async () => {{
            textarea.focus();
            textarea.select();
            try {{
              await navigator.clipboard.writeText(text);
              status.textContent = "Copied. Paste directly into Excel.";
            }} catch (error) {{
              document.execCommand("copy");
              status.textContent = "Selected. Press Ctrl+C if the browser blocked clipboard access.";
            }}
          }});
        }})();
        </script>
        """
        with self.copy_output:
            clear_output(wait=True)
            display(HTML(html_block))


AdaptiveScreenerApp = MarginStabilityScreenerApp
