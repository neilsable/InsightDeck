from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Tuple

import pandas as pd

import matplotlib
matplotlib.use("Agg")  # serverless/headless safe
import matplotlib.pyplot as plt

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


REQUIRED_COLS = {"day", "service", "usage_units", "cost_gbp", "incidents", "sla_pct"}


@dataclass
class KPIs:
    total_cost: float
    total_usage: float
    avg_sla: float
    total_incidents: int


def _validate_and_load(csv_path: Path) -> pd.DataFrame:
    df = pd.read_csv(csv_path)

    missing = REQUIRED_COLS - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}")

    df["day"] = pd.to_datetime(df["day"], errors="coerce")
    if df["day"].isna().any():
        raise ValueError("Invalid date format in 'day'. Use YYYY-MM-DD.")

    # Coerce numeric cols
    for c in ["usage_units", "cost_gbp", "incidents", "sla_pct"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    if df[["usage_units", "cost_gbp", "incidents", "sla_pct"]].isna().any().any():
        raise ValueError("One or more numeric columns contain invalid values.")

    df = df.sort_values("day")
    return df


def _compute_kpis(df: pd.DataFrame) -> KPIs:
    return KPIs(
        total_cost=float(df["cost_gbp"].sum()),
        total_usage=float(df["usage_units"].sum()),
        avg_sla=float(df["sla_pct"].mean()),
        total_incidents=int(df["incidents"].sum()),
    )


def _build_cost_trend_chart(df: pd.DataFrame, out_png: Path) -> None:
    # Aggregate by day (or if multiple services per day)
    daily = df.groupby("day", as_index=False)["cost_gbp"].sum()

    plt.figure(figsize=(8, 3.2))
    plt.plot(daily["day"], daily["cost_gbp"], marker="o")
    plt.title("Cloud Cost Trend (GBP)")
    plt.xlabel("Day")
    plt.ylabel("Cost (GBP)")
    plt.tight_layout()
    plt.savefig(out_png, dpi=160)
    plt.close()


def _add_title(slide, title: str, subtitle: str | None = None) -> None:
    slide.shapes.title.text = title
    title_tf = slide.shapes.title.text_frame
    title_tf.paragraphs[0].font.size = Pt(40)

    if subtitle:
        # Add a subtitle text box
        box = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(11.3), Inches(0.6))
        tf = box.text_frame
        tf.text = subtitle
        tf.paragraphs[0].font.size = Pt(18)
        tf.paragraphs[0].font.color.rgb = RGBColor(90, 90, 90)


def _add_kpi_tile(slide, x, y, w, h, label: str, value: str) -> None:
    shape = slide.shapes.add_shape(1, x, y, w, h)  # 1 = MSO_SHAPE.RECTANGLE
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(245, 247, 250)
    shape.line.color.rgb = RGBColor(210, 215, 222)

    tf = shape.text_frame
    tf.clear()

    p1 = tf.paragraphs[0]
    p1.text = label
    p1.font.size = Pt(14)
    p1.font.color.rgb = RGBColor(80, 80, 80)

    p2 = tf.add_paragraph()
    p2.text = value
    p2.font.size = Pt(24)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(20, 20, 20)


def _format_gbp(x: float) -> str:
    if x >= 1_000_000:
        return f"£{x/1_000_000:.2f}M"
    if x >= 1_000:
        return f"£{x/1_000:.1f}K"
    return f"£{x:.0f}"


def generate_ppt_from_csv(csv_path: Path, out_pptx: Path) -> None:
    csv_path = Path(csv_path)
    out_pptx = Path(out_pptx)

    df = _validate_and_load(csv_path)
    kpis = _compute_kpis(df)

    # Prepare chart image in /tmp
    chart_path = Path("/tmp") / f"insightdeck_cost_trend_{out_pptx.stem}.png"
    _build_cost_trend_chart(df, chart_path)

    prs = Presentation()

    # Slide 1: KPI + Chart
    slide1 = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only
    _add_title(slide1, "InsightDeck Executive Summary", "KPI one-pager + auto-fitted trend chart")

    # KPI tiles row
    left = Inches(1.0)
    top = Inches(2.3)
    tile_w = Inches(2.9)
    tile_h = Inches(1.2)
    gap = Inches(0.35)

    _add_kpi_tile(slide1, left + (tile_w + gap)*0, top, tile_w, tile_h, "Total Cost", _format_gbp(kpis.total_cost))
    _add_kpi_tile(slide1, left + (tile_w + gap)*1, top, tile_w, tile_h, "Total Usage", f"{kpis.total_usage:,.0f}")
    _add_kpi_tile(slide1, left + (tile_w + gap)*2, top, tile_w, tile_h, "Avg SLA", f"{kpis.avg_sla:.2f}%")
    _add_kpi_tile(slide1, left + (tile_w + gap)*3, top, tile_w, tile_h, "Incidents", f"{kpis.total_incidents:,}")

    # Chart
    slide1.shapes.add_picture(str(chart_path), Inches(1.0), Inches(4.0), width=Inches(11.3))

    # Slide 2: Narrative
    slide2 = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only
    _add_title(slide2, "Insights, Risks & Actions", "Narrative appendix (auto-generated)")

    # Build narrative blocks
    # Simple but credible rules (you can refine later)
    daily = df.groupby("day", as_index=False).agg(
        cost_gbp=("cost_gbp", "sum"),
        incidents=("incidents", "sum"),
        sla_pct=("sla_pct", "mean"),
        usage_units=("usage_units", "sum"),
    )
    cost_change = (daily["cost_gbp"].iloc[-1] - daily["cost_gbp"].iloc[0]) if len(daily) > 1 else 0.0
    trend = "up" if cost_change > 0 else "down" if cost_change < 0 else "flat"

    insights = [
        f"Cost trend is {trend} over the period ({_format_gbp(abs(cost_change))} net change).",
        f"Average SLA is {kpis.avg_sla:.2f}%, with {kpis.total_incidents} total incidents.",
        f"Peak daily cost: {_format_gbp(float(daily['cost_gbp'].max()))}.",
    ]
    risks = [
        "Rising cost with flat usage can indicate inefficiency or pricing drift.",
        "Incident spikes may correlate with SLA degradation and operational risk.",
    ]
    actions = [
        "Review top-cost services and apply usage caps / rightsizing opportunities.",
        "Investigate days with incident spikes; add alerts and runbooks.",
        "Set a weekly cost/SLA cadence deck for stakeholders using InsightDeck.",
    ]
    method = [
        "Input validated for required columns and numeric types.",
        "Daily aggregation used for trend chart.",
        "Deck generated server-side using python-pptx; charts rendered via matplotlib (Agg).",
    ]

    # Add text boxes in two columns
    def add_block(title: str, lines: list[str], x: float, y: float) -> None:
        box = slide2.shapes.add_textbox(Inches(x), Inches(y), Inches(5.6), Inches(2.0))
        tf = box.text_frame
        tf.clear()
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.size = Pt(18)
        p0.font.bold = True

        for ln in lines:
            p = tf.add_paragraph()
            p.text = ln
            p.level = 0
            p.font.size = Pt(13)

    add_block("Key Insights", insights, 1.0, 2.2)
    add_block("Key Risks", risks, 7.0, 2.2)
    add_block("Recommended Actions", actions, 1.0, 4.6)
    add_block("Method Notes", method, 7.0, 4.6)

    prs.save(out_pptx)
