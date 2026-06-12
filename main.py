import sys
import subprocess
import plotly.graph_objects as go
from dash import Dash, dcc, html, dash_table, Input, Output, callback_context, no_update

def ensure_package(pkg_name, import_name=None):
    try:
        __import__(import_name or pkg_name)
    except ImportError:
        print(f"Installing {pkg_name} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg_name])

# Ensure required packages are installed
ensure_package("pandas")
ensure_package("dash")
ensure_package("plotly")
ensure_package("openpyxl")

import pandas as pd
from dash import Dash, dcc, html, dash_table, Input, Output
import plotly.express as px

# ---------- Constants for styling ----------

ACCENT_COLOR = "#1976D2"
PRIMARY_TEXT = "#37474F"
BACKGROUND = "#F5F5F5"

TITLE_FONT_SIZE = 26
SECTION_TITLE_SIZE = 20
AXIS_TITLE_SIZE = 14
AXIS_TICK_SIZE = 11

TAB_STYLE = {
    "fontWeight": "bold",
    "fontSize": "15px",
    "padding": "10px 20px",
    "border": "none",
}

TAB_SELECTED_STYLE = {
    "fontWeight": "bold",
    "fontSize": "15px",
    "padding": "10px 20px",
    "border": "none",
    "borderBottom": f"3px solid {ACCENT_COLOR}",
    "color": ACCENT_COLOR,
    "backgroundColor": "#E3F2FD",
}

# ---------- Load & prepare data ----------

excel_path = "Food Analysis Data.xlsx"

event_df = pd.read_excel(excel_path, sheet_name="Event")
caterer_df = pd.read_excel(excel_path, sheet_name="Caterer")
food_df = pd.read_excel(excel_path, sheet_name="FoodOrders")
event_orders_df = pd.read_excel(excel_path, sheet_name="EventOrders")

def smart_title(s: str) -> str:
    # Keep text that is fully upper-case (pure acronyms etc.),
    # otherwise use Title Case.
    return s if s.isupper() else s.title()

# Title-case everything first
for df in (event_df, caterer_df, food_df):
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).apply(smart_title)

# Fix DSAC acronym so it stays capitalized in event names
if "event_name" in event_df.columns:
    event_df["event_name"] = (
        event_df["event_name"]
        .str.replace("Dsac", "DSAC", regex=False)
        .str.replace("Dsca", "DSAC", regex=False)  # just in case old spelling exists
    )

# --- Fix specific Caterer name capitalization issues ---
if "caterer_name" in caterer_df.columns:
    caterer_df["caterer_name"] = (
        caterer_df["caterer_name"]
        .str.replace("Dibella'S Subs", "DiBella's Subs", regex=False)
        .str.replace("Bj'S", "BJ's", regex=False)
        .str.replace("Rit", "RIT", regex=False)
        .str.replace("Moe'S", "Moe's", regex=False)
        .str.replace("Lugia'S Ice Cream", "Lugia's Ice Cream", regex=False)
    )

# Join everything into one wide table
full_df = (
    event_orders_df
    .merge(event_df, on="event_id", how="left")
    .merge(caterer_df, on="caterer_id", how="left")
    .merge(food_df, on="food_order_id", how="left", suffixes=("_event", "_food"))
)

# Names of the DSAC luncheons (normalized to lowercase for easy matching)
DSAC_EVENTS_NORMALIZED = {
    "dsac fall luncheon",
    "dsac spring luncheon",
}

# Make sure numeric columns are numeric
for col in ["cost", "total_unit_cost", "quantity", "leftovers",
            "unit_price", "fee", "attendees_num", "rsvp_num"]:
    if col in full_df.columns:
        full_df[col] = pd.to_numeric(full_df[col], errors="coerce")

# ---- Aggregated tables ----

# Event-level metrics
event_metrics = (
    full_df
    .groupby(["event_id", "event_name", "attendees_num", "rsvp_num"], as_index=False)
    .agg(
        total_cost=("cost", "sum"),
        total_quantity=("quantity", "sum"),
        total_leftovers=("leftovers", "sum"),
    )
)
event_metrics["cost_per_attendee"] = (
    event_metrics["total_cost"] / event_metrics["attendees_num"]
)

# Caterer-level spend
caterer_metrics = (
    full_df
    .groupby("caterer_name", dropna=False, as_index=False)
    .agg(total_cost=("cost", "sum"))
    .sort_values("total_cost", ascending=False)
)
caterer_metrics["caterer_name"] = caterer_metrics["caterer_name"].fillna(
    "Unknown / Not Recorded"
)

# Food-level popularity & waste
food_metrics = (
    full_df
    .groupby("order", as_index=False)
    .agg(
        total_quantity=("quantity", "sum"),
        total_cost=("cost", "sum"),
        avg_leftovers=("leftovers", "mean"),
    )
    .sort_values("total_quantity", ascending=False)
)

# ---------- Options for Food Forecast tab ----------

CATERER_OPTIONS = [
    {"label": name, "value": name}
    for name in sorted(full_df["caterer_name"].dropna().unique())
]

FOOD_OPTIONS = [
    {"label": name, "value": name}
    for name in sorted(full_df["order"].dropna().unique())
]

# ---------- Dropdown options (remove DSAC Fall Luncheon) ----------

def build_event_dropdown_options():
    options = [{
        "label": "DSAC Fall/Spring Luncheon (Combined)",
        "value": "dsac_combined",
    }]

    for _, row in event_metrics.iterrows():
        name = str(row["event_name"])
        lower = name.lower()

        # Hide ANY DSAC luncheon (Fall or Spring)
        if "dsac" in lower and "luncheon" in lower:
            continue

        options.append({"label": name, "value": row["event_id"]})

    return options

EVENT_DROPDOWN_OPTIONS = build_event_dropdown_options()

FORECAST_EVENT_OPTIONS = (
    [{"label": "Custom / New Event", "value": "custom"}] +
    [
        {"label": name, "value": eid}
        for eid, name in event_metrics[["event_id", "event_name"]].values
        # NO DSAC filtering here → Fall & Spring both appear
    ]
)

# Defaults for the budget-forecast tab
if "cost_per_attendee" in event_metrics.columns and not event_metrics["cost_per_attendee"].isna().all():
    DEFAULT_COST_PER_ATTENDEE = float(event_metrics["cost_per_attendee"].mean())
else:
    DEFAULT_COST_PER_ATTENDEE = 20.0  # fallback

if "attendees_num" in event_metrics.columns and not event_metrics["attendees_num"].isna().all():
    DEFAULT_ATTENDEES = int(event_metrics["attendees_num"].mean())
else:
    DEFAULT_ATTENDEES = 50  # fallback

# ---------- Helper functions for figures & formatting ----------

def style_bar(fig):
    fig.update_layout(
        title_font=dict(size=SECTION_TITLE_SIZE, color=ACCENT_COLOR),
        xaxis_title_font=dict(size=AXIS_TITLE_SIZE, color=PRIMARY_TEXT),
        yaxis_title_font=dict(size=AXIS_TITLE_SIZE, color=PRIMARY_TEXT),
        xaxis_tickfont=dict(size=AXIS_TICK_SIZE, color=PRIMARY_TEXT),
        yaxis_tickfont=dict(size=AXIS_TICK_SIZE, color=PRIMARY_TEXT),
        bargap=0.25,
        margin=dict(t=70, b=90, l=70, r=40),
        plot_bgcolor="white",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    fig.update_xaxes(tickangle=-40, ticks="outside")
    fig.update_yaxes(tickformat=",.2f")
    fig.update_traces(hoverlabel=dict(font_size=12))
    return fig

def style_scatter(fig):
    fig.update_layout(
        title_font=dict(size=SECTION_TITLE_SIZE, color=ACCENT_COLOR),
        xaxis_title_font=dict(size=AXIS_TITLE_SIZE, color=PRIMARY_TEXT),
        yaxis_title_font=dict(size=AXIS_TITLE_SIZE, color=PRIMARY_TEXT),
        xaxis_tickfont=dict(size=AXIS_TICK_SIZE, color=PRIMARY_TEXT),
        yaxis_tickfont=dict(size=AXIS_TICK_SIZE, color=PRIMARY_TEXT),
        margin=dict(t=70, b=80, l=70, r=40),
        plot_bgcolor="white",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    fig.update_xaxes(ticks="outside")
    fig.update_yaxes(tickformat=",.2f")
    fig.update_traces(
        marker=dict(size=10, line=dict(width=1, color="#888")),
        hoverlabel=dict(font_size=12),
    )
    return fig

def format_currency(value):
    if pd.isna(value):
        return ""
    return f"${value:,.2f}"

# ---------- Build the Dash app ----------

app = Dash(__name__)
app.title = "Food Analysis Dashboard"

app.layout = html.Div(
    style={
        "fontFamily": "Arial, sans-serif",
        "backgroundColor": BACKGROUND,
        "minHeight": "100vh",
        "padding": "40px 20px 60px 20px",
    },
    children=[
        html.Div(
            style={"maxWidth": "1200px", "margin": "0 auto"},
            children=[
                html.H1(
                    "Food Analysis Dashboard",
                    style={
                        "textAlign": "center",
                        "color": PRIMARY_TEXT,
                        "fontSize": TITLE_FONT_SIZE,
                        "fontWeight": "bold",
                        "marginBottom": "30px",
                        "marginTop": "10px",
                    },
                ),

                # Summary cards
                html.Div(
                    style={
                        "display": "flex",
                        "gap": "20px",
                        "flexWrap": "wrap",
                        "justifyContent": "center",
                        "marginBottom": "30px",
                    },
                    children=[
                        # ---- Total Events Card ----
                        html.Div(
                            style={
                                "border": "1px solid #ddd",
                                "padding": "10px 15px",
                                "borderRadius": "10px",
                                "backgroundColor": "white",
                                "minWidth": "180px",
                                "maxWidth": "220px",
                                "height": "140px",
                                "display": "flex",
                                "flexDirection": "column",
                                "justifyContent": "space-evenly",
                                "alignItems": "center",
                                "boxShadow": "0 1px 3px rgba(0,0,0,0.1)",
                                "textAlign": "center",
                            },
                            children=[
                                html.H4(
                                    "Total Events",
                                    style={
                                        "margin": "0",
                                        "fontWeight": "bold",
                                        "fontSize": "18px",
                                    },
                                ),
                                html.P(
                                    f"{event_metrics['event_id'].nunique()}",
                                    style={
                                        "fontSize": "22px",
                                        "fontWeight": "bold",
                                        "margin": "0",
                                    },
                                ),
                            ],
                        ),

                        # ---- Total Spend Card ----
                        html.Div(
                            style={
                                "border": "1px solid #ddd",
                                "padding": "10px 15px",
                                "borderRadius": "10px",
                                "backgroundColor": "white",
                                "minWidth": "180px",
                                "maxWidth": "220px",
                                "height": "140px",
                                "display": "flex",
                                "flexDirection": "column",
                                "justifyContent": "space-evenly",
                                "alignItems": "center",
                                "boxShadow": "0 1px 3px rgba(0,0,0,0.1)",
                                "textAlign": "center",
                            },
                            children=[
                                html.H4(
                                    "Total Spend",
                                    style={
                                        "margin": "0",
                                        "fontWeight": "bold",
                                        "fontSize": "18px",
                                    },
                                ),
                                html.P(
                                    format_currency(event_metrics["total_cost"].sum()),
                                    style={
                                        "fontSize": "22px",
                                        "fontWeight": "bold",
                                        "color": ACCENT_COLOR,
                                        "margin": "0",
                                    },
                                ),
                            ],
                        ),

                        # ---- Avg Cost Per Attendee Card ----
                        html.Div(
                            style={
                                "border": "1px solid #ddd",
                                "padding": "10px 15px",
                                "borderRadius": "10px",
                                "backgroundColor": "white",
                                "minWidth": "160px",
                                "maxWidth": "200px",
                                "height": "140px",
                                "display": "flex",
                                "flexDirection": "column",
                                "justifyContent": "space-evenly",
                                "alignItems": "center",
                                "boxShadow": "0 1px 3px rgba(0,0,0,0.1)",
                                "textAlign": "center",
                            },
                            children=[
                                html.H4(
                                    "Average Cost Per Attendee",
                                    style={
                                        "marginTop": "20px",      # <-- move the title downward
                                        "marginBottom": "10px",
                                        "fontWeight": "bold",
                                        "fontSize": "18px",
                                    },
                                ),
                                html.P(
                                    format_currency(event_metrics["cost_per_attendee"].mean()),
                                    style={
                                        "fontSize": "22px",
                                        "fontWeight": "bold",
                                        "color": ACCENT_COLOR,
                                        "marginTop": "0",
                                    },
                                ),
                            ],
                        ),
                    ],
                ),


dcc.Tabs(
    style={"marginTop": "10px"},
    children=[
        # -------- Event Overview Tab --------
        dcc.Tab(
            label="Event Overview",
            style=TAB_STYLE,
            selected_style=TAB_SELECTED_STYLE,
            children=[
                html.Br(),
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "20px",
                        "justifyContent": "space-between",
                        "marginBottom": "30px",
                    },
                    children=[
                        html.Div(
                            style={"width": "100%", "maxWidth": "580px"},
                            children=[
                                dcc.Graph(
                                    id="event-cost-bar",
                                    figure=style_bar(
                                        px.bar(
                                            event_metrics,
                                            x="event_name",
                                            y="total_cost",
                                            title="Total Cost Per Event",
                                            labels={
                                                "event_name": "Event",
                                                "total_cost": "Total Cost ($)",
                                            },
                                        )
                                    ),
                                )
                            ],
                        ),
                        html.Div(
                            style={"width": "100%", "maxWidth": "580px"},
                            children=[
                                dcc.Graph(
                                    id="cost-vs-attendees",
                                    figure=style_scatter(
                                        px.scatter(
                                            event_metrics,
                                            x="attendees_num",
                                            y="cost_per_attendee",
                                            size="total_cost",
                                            size_max=40,
                                            hover_name="event_name",
                                            title="Cost Per Attendee vs. Event Size",
                                            labels={
                                                "attendees_num": "Number Of Attendees",
                                                "cost_per_attendee": "Cost Per Attendee ($)",
                                            },
                                        )
                                    ),
                                )
                            ],
                        ),
                    ],
                ),
            ],
        ),

        # -------- Caterers Tab --------
        dcc.Tab(
            label="Caterers",
            style=TAB_STYLE,
            selected_style=TAB_SELECTED_STYLE,
            children=[
                html.Br(),
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "20px",
                        "justifyContent": "space-between",
                        "marginBottom": "30px",
                    },
                    children=[
                        html.Div(
                            style={"width": "100%", "maxWidth": "580px"},
                            children=[
                                dcc.Graph(
                                    id="caterer-cost-bar",
                                    figure=style_bar(
                                        px.bar(
                                            caterer_metrics,
                                            x="caterer_name",
                                            y="total_cost",
                                            title="Total Spend By Caterer",
                                            labels={
                                                "caterer_name": "Caterer",
                                                "total_cost": "Total Cost ($)",
                                            },
                                        )
                                    ),
                                )
                            ],
                        ),
                        html.Div(
                            style={
                                "width": "100%",
                                "maxWidth": "580px",
                                "margin": "0 auto",
                            },
                            children=[
                                html.H3(
                                    "Caterer Spend Table",
                                    style={
                                        "textAlign": "center",
                                        "fontSize": SECTION_TITLE_SIZE,
                                        "fontWeight": "bold",
                                        "color": PRIMARY_TEXT,
                                        "marginBottom": "10px",
                                    },
                                ),
                                dash_table.DataTable(
                                    id="caterer-table",
                                    data=pd.DataFrame({
                                        "Caterer": caterer_metrics["caterer_name"],
                                        "Total Cost": caterer_metrics["total_cost"].map(
                                            format_currency
                                        ),
                                    }).to_dict("records"),
                                    columns=[
                                        {"name": "Caterer", "id": "Caterer"},
                                        {"name": "Total Cost", "id": "Total Cost"},
                                    ],
                                    page_size=10,
                                    style_table={
                                        "maxWidth": "100%",
                                        "margin": "0 auto",
                                        "border": "1px solid #ddd",
                                        "borderRadius": "8px",
                                        "overflowX": "auto",
                                    },
                                    style_cell={
                                        "textAlign": "left",
                                        "padding": "8px",
                                        "fontSize": 13,
                                    },
                                    style_header={
                                        "fontWeight": "bold",
                                        "backgroundColor": "#ECEFF1",
                                    },
                                ),
                            ],
                        ),
                    ],
                ),
            ],
        ),

        # -------- Food Items Tab --------
        dcc.Tab(
            label="Food Items",
            style=TAB_STYLE,
            selected_style=TAB_SELECTED_STYLE,
            children=[
                html.Br(),
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "20px",
                        "justifyContent": "space-between",
                        "marginBottom": "30px",
                    },
                    children=[
                        html.Div(
                            style={"width": "100%", "maxWidth": "580px"},
                            children=[
                                dcc.Graph(
                                    id="food-quantity-bar",
                                    figure=style_bar(
                                        px.bar(
                                            food_metrics.head(15),
                                            x="order",
                                            y="total_quantity",
                                            title="Top Food Items By Quantity Ordered",
                                            labels={
                                                "order": "Food Item",
                                                "total_quantity": "Total Quantity",
                                            },
                                        )
                                    ),
                                )
                            ],
                        ),
                        html.Div(
                            style={"width": "100%", "maxWidth": "580px"},
                            children=[
                                dcc.Graph(
                                    id="food-leftovers-scatter",
                                    figure=style_scatter(
                                        px.scatter(
                                            food_metrics.assign(
                                                leftover_rate_pct=(
                                                    food_metrics["avg_leftovers"]
                                                    .div(food_metrics["total_quantity"])
                                                    .fillna(0)
                                                    * 100
                                                ).round(1),
                                                total_cost_formatted=food_metrics[
                                                    "total_cost"
                                                ].map(format_currency),
                                            ),
                                            x="total_quantity",
                                            y="avg_leftovers",
                                            hover_name="order",
                                            hover_data={
                                                "total_quantity": True,
                                                "avg_leftovers": True,
                                                "leftover_rate_pct": True,
                                                "total_cost_formatted": True,
                                            },
                                            title="Food Quantity vs. Average Leftovers",
                                            labels={
                                                "total_quantity": "Total Quantity Ordered",
                                                "avg_leftovers": "Average Leftovers",
                                                "leftover_rate_pct": "Leftover Rate (%)",
                                                "total_cost_formatted": "Total Cost",
                                            },
                                        )
                                    ),
                                )
                            ],
                        ),
                    ],
                ),
            ],
        ),

        # -------- Event Detail Tab --------
        dcc.Tab(
            label="Event Detail",
            style=TAB_STYLE,
            selected_style=TAB_SELECTED_STYLE,
            children=[
                html.Br(),
                html.Div(
                    style={
                        "display": "flex",
                        "flexDirection": "column",
                        "alignItems": "flex-start",
                        "marginBottom": "20px",
                    },
                    children=[
                        html.Label(
                            "Select An Event:",
                            style={
                                "fontWeight": "bold",
                                "marginBottom": "8px",
                                "fontSize": 14,
                            },
                        ),
                        dcc.Dropdown(
                            id="event-dropdown",
                            options=(
                                    [
                                        {
                                            "label": "DSAC Fall/Spring Luncheon (Combined)",
                                            "value": "dsac_combined",  # <- use this spelling everywhere
                                        }
                                    ]
                                    + [
                                        {"label": name, "value": eid}
                                        for eid, name in
                                        event_metrics[["event_id", "event_name"]].values
                                        # remove the individual DSAC Fall/Spring luncheons from the dropdown
                                        if name.strip().lower() not in DSAC_EVENTS_NORMALIZED
                                    ]
                            ),
                            value="dsac_combined",
                            multi=False,
                            style={
                                "width": "350px",
                                "border": "2px solid #1976D2",
                                "borderRadius": "8px",
                                "backgroundColor": "#E3F2FD",
                                "color": "#0D47A1",
                                "fontWeight": "500",
                                "padding": "4x 8x",
                            }
                        ),
                    ],
                ),
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "20px",
                        "justifyContent": "space-between",
                        "marginBottom": "30px",
                    },
                    children=[
                        html.Div(
                            style={"width": "100%", "maxWidth": "580px"},
                            children=[
                                dcc.Graph(id="event-food-breakdown"),
                            ],
                        ),
                        html.Div(
                            style={
                                "width": "100%",
                                "maxWidth": "580px",
                                "margin": "0 auto",
                            },
                            children=[
                                html.H3(
                                    "Event Food Detail Table",
                                    style={
                                        "textAlign": "center",
                                        "fontSize": SECTION_TITLE_SIZE,
                                        "fontWeight": "bold",
                                        "color": PRIMARY_TEXT,
                                        "marginBottom": "10px",
                                    },
                                ),
                                dash_table.DataTable(
                                    id="event-detail-table",
                                    page_size=10,
                                    style_table={
                                        "maxWidth": "100%",
                                        "margin": "0 auto",
                                        "border": "1px solid #ddd",
                                        "borderRadius": "8px",
                                        "overflowX": "auto",
                                    },
                                    style_cell={
                                        "textAlign": "left",
                                        "padding": "8px",
                                        "fontSize": 13,
                                    },
                                    style_header={
                                        "fontWeight": "bold",
                                        "backgroundColor": "#ECEFF1",
                                    },
                                ),
                            ],
                        ),
                    ],
                ),
            ],
        ),

        # -------- Budget Forecast Tab --------
        dcc.Tab(
            label="Budget Forecast",
            style=TAB_STYLE,
            selected_style=TAB_SELECTED_STYLE,
            children=[
                html.Br(),
                html.Div(
                    style={
                        "maxWidth": "600px",
                        "margin": "0 auto",
                        "display": "flex",
                        "flexDirection": "column",
                        "gap": "12px",
                    },
                    children=[
                        # ---- Base Event dropdown ----
                        html.Label(
                            "Base Event (optional):",
                            style={"fontWeight": "bold", "marginBottom": "6px"},
                        ),
                        dcc.Dropdown(
                            id="forecast-event-dropdown",
                            options=FORECAST_EVENT_OPTIONS,
                            value="custom",
                            clearable=False,
                            style={
                                "width": "350px",
                                "border": "1px solid #1976D2",
                                "borderRadius": "8px",
                                "backgroundColor": "#E3F2FD",
                                "color": "#0D47A1",
                                "fontWeight": "500",
                                "marginBottom": "18px",
                            },
                        ),

                        # ---- Planned attendees ----
                        html.Label(
                            "Planned Number of Attendees:",
                            style={"fontWeight": "bold"},
                        ),
                        dcc.Input(
                            id="forecast-attendees",
                            type="number",
                            value=50,
                            min=0,
                            style={"width": "160px"},
                        ),

                        # ---- Cost per attendee ----
                        html.Label(
                            "Estimated Cost Per Attendee ($):",
                            style={"fontWeight": "bold", "marginTop": "10px"},
                        ),
                        dcc.Input(
                            id="forecast-cost-per-attendee",
                            type="number",
                            value=20,
                            min=0,
                            step=0.01,
                            style={"width": "160px"},
                        ),

                        # ---- Budget text ----
                        html.H3(
                            id="forecast-budget-value",
                            style={
                                "marginTop": "20px",
                                "color": ACCENT_COLOR,
                                "fontWeight": "bold",
                            },
                        ),

                        # ---- Suggested quantities table ----
                        html.H4(
                            "Suggested Quantities by Food Item",
                            style={
                                "marginTop": "10px",
                                "fontWeight": "bold",
                                "color": PRIMARY_TEXT,
                            },
                        ),
                        dash_table.DataTable(
                            id="forecast-food-table",
                            columns=[
                                {"name": "Food Item", "id": "Food Item"},
                                {"name": "Expected Qty", "id": "Expected Qty"},
                            ],
                            data=[],
                            page_size=10,
                            style_table={
                                "maxWidth": "100%",
                                "margin": "0 auto",
                                "border": "1px solid #ddd",
                                "borderRadius": "8px",
                                "overflowX": "auto",
                            },
                            style_cell={
                                "textAlign": "left",
                                "padding": "8px",
                                "fontSize": 13,
                            },
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#ECEFF1",
                            },
                        ),
                    ],
                ),
            ],
        ),

# -------- Food Forecast Tab --------
                dcc.Tab(
                    label="Food Forecast",
                    style=TAB_STYLE,
                    selected_style=TAB_SELECTED_STYLE,
                    children=[
                        html.Br(),
                        html.Div(
                            style={
                                "maxWidth": "650px",
                                "margin": "0 auto",
                                "display": "flex",
                                "flexDirection": "column",
                                "gap": "14px",
                            },
                            children=[
                                html.H3(
                                    "Food Quantity Forecast",
                                    style={
                                        "fontWeight": "bold",
                                        "color": PRIMARY_TEXT,
                                        "marginBottom": "4px",
                                    },
                                ),

                                html.Label(
                                    "Caterer:",
                                    style={"fontWeight": "bold"},
                                ),
                                dcc.Dropdown(
                                    id="ff-caterer",
                                    options=CATERER_OPTIONS,
                                    placeholder="Select a caterer...",
                                    style={
                                        "width": "100%",
                                        "border": "1px solid #1976D2",
                                        "borderRadius": "8px",
                                        "backgroundColor": "#E3F2FD",
                                    },
                                ),

                                html.Label(
                                    "Food Item:",
                                    style={"fontWeight": "bold", "marginTop": "6px"},
                                ),
                                dcc.Dropdown(
                                    id="ff-food",
                                    options=FOOD_OPTIONS,
                                    placeholder="Select a food item...",
                                    style={
                                        "width": "100%",
                                        "border": "1px solid #1976D2",
                                        "borderRadius": "8px",
                                        "backgroundColor": "#E3F2FD",
                                    },
                                ),

                                html.Label(
                                    "Expected Number of Attendees:",
                                    style={"fontWeight": "bold", "marginTop": "6px"},
                                ),
                                dcc.Input(
                                    id="ff-attendees",
                                    type="number",
                                    value=50,
                                    min=1,
                                    step=1,
                                    style={"width": "160px"},
                                ),

                                html.Div(
                                    id="food-forecast-output",
                                    style={
                                        "marginTop": "18px",
                                        "padding": "12px 14px",
                                        "borderRadius": "8px",
                                        "backgroundColor": "white",
                                        "border": "1px solid #ddd",
                                    },
                                ),
                            ],
                        ),
                    ],
                ),
    ],

),

                html.Hr(style={"marginTop": "40px"}),
                html.Footer(
                    "Internal Food Analysis Dashboard – For Planning Use Only",
                    style={
                        "textAlign": "center",
                        "fontSize": 12,
                        "color": "#777",
                        "marginTop": "10px",
                    },
                ),
            ],
        ),
    ],
)

# ---------- Callbacks ----------

@app.callback(
    Output("event-food-breakdown", "figure"),
    Output("event-detail-table", "data"),
    Output("event-detail-table", "columns"),
    Input("event-dropdown", "value"),
)
def update_event_detail(selected_event_id):
    # ----- choose which rows to use -----
    if selected_event_id == "dsac_combined":
        # combine DSAC Fall + Spring luncheons using normalized names
        mask = full_df["event_name"].str.strip().str.lower().isin(DSAC_EVENTS_NORMALIZED)
        event_rows = full_df[mask].copy()
    else:
        event_rows = full_df[full_df["event_id"] == selected_event_id].copy()

    if event_rows.empty:
        empty_fig = style_bar(px.bar(title="No data available for the selected event(s)"))
        return empty_fig, [], []

    # ----- consumed vs leftovers per food item -----
    food_breakdown = (
        event_rows
        .groupby("order", as_index=False)
        .agg(
            quantity=("quantity", "sum"),
            leftovers=("leftovers", "sum"),
            total_cost=("cost", "sum"),
        )
    )

    food_breakdown["leftovers"] = food_breakdown["leftovers"].fillna(0)
    food_breakdown["consumed"] = food_breakdown["quantity"] - food_breakdown["leftovers"]

    food_long = food_breakdown.melt(
        id_vars="order",
        value_vars=["consumed", "leftovers"],
        var_name="Status",
        value_name="Units",
    )

    if selected_event_id == "dsac_combined":
        chart_title = "DSAC Fall/Spring Luncheon"
    else:
        event_name = str(event_rows["event_name"].iloc[0])
        chart_title = f"{event_name}"

    fig = px.bar(
        food_long,
        x="order",
        y="Units",
        color="Status",
        barmode="stack",
        title=chart_title,
        labels={"order": "Food Item", "Units": "Units", "Status": "Status"},
    )
    fig = style_bar(fig)
    fig.update_yaxes(tickprefix="")

    # ----- detail table -----
    display_cols = [
        "event_name", "caterer_name", "order", "dietary_label",
        "quantity", "leftovers", "unit_price", "fee", "cost",
    ]
    existing_cols = [c for c in display_cols if c in event_rows.columns]
    table_df = event_rows[existing_cols].copy()

    rename_map = {
        "event_name": "Event",
        "caterer_name": "Caterer",
        "order": "Food Item",
        "dietary_label": "Dietary Label",
        "quantity": "Quantity",
        "leftovers": "Leftovers",
        "unit_price": "Unit Price",
        "fee": "Fee",
        "cost": "Total Cost",
    }
    table_df = table_df.rename(columns=rename_map)

    if "Dietary Label" in table_df.columns:
        table_df["Dietary Label"] = (
            table_df["Dietary Label"]
            .astype(str)
            .replace("Nan", "N/A")
            .replace("None", "N/A")
        )

    for col in ["Unit Price", "Fee", "Total Cost"]:
        if col in table_df.columns:
            table_df[col] = table_df[col].map(format_currency)

    sort_cols = [c for c in ["Event", "Food Item"] if c in table_df.columns]
    if sort_cols:
        table_df = table_df.sort_values(sort_cols)

    data = table_df.to_dict("records")
    columns = [{"name": col, "id": col} for col in table_df.columns]

    return fig, data, columns

# 1) Keep attendees + cost in sync based on base event and attendee changes
@app.callback(
    Output("forecast-attendees", "value"),
    Output("forecast-cost-per-attendee", "value"),
    Input("forecast-event-dropdown", "value"),
    Input("forecast-attendees", "value"),
    Input("forecast-cost-per-attendee", "value"),
)
def sync_attendees_and_cost(selected_event, attendees, cost_input):
    ctx = callback_context

    # Use the raw values Dash gives us; don't auto-replace None yet
    attendees_val = attendees
    cost_val = cost_input

    # ---------- Initial load: nothing triggered yet ----------
    if not ctx.triggered:
        # Just show your defaults on first render
        return DEFAULT_ATTENDEES, round(DEFAULT_COST_PER_ATTENDEE, 2)

    trigger_id = ctx.triggered[0]["prop_id"].split(".")[0]

    # ---------- 1) User manually edited COST ----------
    if trigger_id == "forecast-cost-per-attendee":
        # Respect what the user typed; don't change it
        # If attendees is None here (user is clearing), leave it alone as well
        return attendees_val, cost_val

    # ---------- 2) Base event dropdown changed ----------
    if trigger_id == "forecast-event-dropdown":
        # Custom / New event → reset to defaults
        if selected_event in (None, "custom"):
            return DEFAULT_ATTENDEES, round(DEFAULT_COST_PER_ATTENDEE, 2)

        # Any single real event (including DSAC Fall & DSAC Spring)
        row = event_metrics.loc[event_metrics["event_id"] == selected_event]
        if row.empty:
            return DEFAULT_ATTENDEES, round(DEFAULT_COST_PER_ATTENDEE, 2)

        row = row.iloc[0]
        attendees_val = int(row["attendees_num"])
        cost_val = float(row["cost_per_attendee"])
        return attendees_val, round(cost_val, 2)

    # ---------- 3) Attendees input changed ----------
    # If user just cleared the box (attendees is None), don't fight them.
    if attendees_val is None:
        return no_update, no_update

    # Custom event: user controls the cost, so don't recompute it
    if selected_event in (None, "custom"):
        return attendees_val, cost_val

    # Single event: recompute cost = total_cost / new_attendees
    row = event_metrics.loc[event_metrics["event_id"] == selected_event]
    if row.empty:
        return attendees_val, cost_val

    total_cost = float(row["total_cost"].iloc[0])
    new_cost = total_cost / max(attendees_val, 1)
    return attendees_val, round(new_cost, 2)


# 2) Build budget text + suggested-quantities table
@app.callback(
    Output("forecast-budget-value", "children"),
    Output("forecast-food-table", "data"),
    Input("forecast-attendees", "value"),
    Input("forecast-cost-per-attendee", "value"),
)
def update_forecast(attendees, cost_per_attendee):
    if not attendees or cost_per_attendee in (None, ""):
        return "Enter values above to see an estimated budget.", []

    budget = attendees * cost_per_attendee

    total_qty_sum = food_metrics["total_quantity"].sum()
    if total_qty_sum == 0:
        table_data = []
    else:
        share = food_metrics["total_quantity"] / total_qty_sum
        expected_qty = (share * attendees).round(0).astype(int)

        table_data = [
            {"Food Item": item, "Expected Qty": int(qty)}
            for item, qty in zip(food_metrics["order"], expected_qty)
        ]

    budget_text = f"Estimated Budget: {format_currency(budget)}"
    return budget_text, table_data

def run_forecast(caterer, food, attendees):
    if not caterer or not food or not attendees:
        return "Please select all forecasting options."

    result = forecast_food_range(df, caterer, food, attendees)

    if "error" in result:
        return result["error"]

    return html.Div([
        html.H4("Forecasted Quantity Range"),
        html.P(f"Recommended Minimum: {result['min_estimate']}"),
        html.P(f"Recommended Maximum: {result['max_estimate']}"),
        html.P(f"Avg Quantity per Person: {result['avg_per_person']}")
    ])

def forecast_food_range(df, caterer_name, food_name, expected_attendees):
    """
    Forecast expected quantity range for a given caterer + food item
    using historical event consumption patterns.

    Returns a dictionary:
    {
        "min_estimate": int,
        "max_estimate": int,
        "avg_per_person": float
    }
    """

    # Filter to this caterer + food
    sub = df[
        (df["caterer_name"] == caterer_name) &
        (df["order"] == food_name)
    ]

    if sub.empty:
        return {
            "error": f"No historical data found for {caterer_name} → {food_name}"
        }

    # Prevent divide-by-zero errors
    sub = sub[sub["attendees_num"] > 0]

    if sub.empty:
        return {
            "error": f"No valid attendee data for {caterer_name} → {food_name}"
        }

    # Compute quantity per person historically
    sub["qty_per_person"] = sub["quantity"] / sub["attendees_num"]

    avg_per_person = sub["qty_per_person"].mean()

    # Forecast values (ensure whole numbers)
    min_estimate = int(max(1, round(avg_per_person * expected_attendees * 0.9)))
    max_estimate = int(max(1, round(avg_per_person * expected_attendees * 1.1)))

    return {
        "min_estimate": min_estimate,
        "max_estimate": max_estimate,
        "avg_per_person": round(avg_per_person, 3)
    }

@app.callback(
    Output("food-forecast-output", "children"),
    Input("ff-caterer", "value"),
    Input("ff-food", "value"),
    Input("ff-attendees", "value"),
)
def update_food_forecast(caterer_name, food_name, expected_attendees):
    # Basic validation
    if not caterer_name or not food_name:
        return html.P(
            "Please select a caterer and a food item.",
            style={"color": "#777"}
        )

    if not expected_attendees or expected_attendees <= 0:
        return html.P(
            "Please enter a positive number of attendees.",
            style={"color": "#777"}
        )

    # Filter data for this caterer + food
    sub = full_df[
        (full_df["caterer_name"] == caterer_name) &
        (full_df["order"] == food_name)
    ]

    if sub.empty:
        return html.P(
            f"No historical data found for {caterer_name} – {food_name}.",
            style={"color": "#B71C1C"}
        )

    # Remove rows with invalid attendee counts
    sub = sub[sub["attendees_num"] > 0]
    if sub.empty:
        return html.P(
            f"No valid attendee data for {caterer_name} – {food_name}.",
            style={"color": "#B71C1C"}
        )

    # Quantity per person historically
    sub = sub.assign(
        qty_per_person=sub["quantity"] / sub["attendees_num"]
    )
    avg_per_person = sub["qty_per_person"].mean()

    # Forecast range (±10%)
    min_estimate = int(max(1, round(avg_per_person * expected_attendees * 0.9)))
    max_estimate = int(max(1, round(avg_per_person * expected_attendees * 1.1)))

    if avg_per_person <= 0:
        people_per_unit_text = "Not enough data to interpret."
    else:
        people_per_unit = round(1 / avg_per_person)
        people_per_unit_text = f"≈ 1 unit per {people_per_unit} people"

    return html.Div([
        html.H4(
            "Forecasted Quantity Range",
            style={"marginBottom": "15px"}
        ),

        html.P(
            f"Recommended Minimum: {min_estimate}",
            style={"marginBottom": "8px", "fontSize": "16px"}
        ),

        html.P(
            f"Recommended Maximum: {max_estimate}",
            style={"marginBottom": "12px", "fontSize": "16px"}
        ),

        html.P(
            f"Average Quantity per Person (historical): {avg_per_person:.2f} "
            f"({people_per_unit_text})",
            style={"marginTop": "10px", "color": "#555", "fontSize": "15px"}
        )
    ])

# ---------- Main ----------

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8050, debug=False)
