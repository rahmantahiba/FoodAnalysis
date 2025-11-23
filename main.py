import sys
import subprocess

def ensure_package(pkg_name, import_name=None):
    """
    Try to import a package; if it's not installed, install it with pip.
    pkg_name: name used with pip (e.g. 'pandas')
    import_name: name used in import (e.g. 'dash', 'plotly.express')
    """
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

# Excel file should be in the same folder as this script
excel_path = "Food Analysis Data.xlsx"

event_df = pd.read_excel(excel_path, sheet_name="Event")
caterer_df = pd.read_excel(excel_path, sheet_name="Caterer")
food_df = pd.read_excel(excel_path, sheet_name="FoodOrders")
event_orders_df = pd.read_excel(excel_path, sheet_name="EventOrders")

# Normalize some strings for nicer display (Title Case)
for df in (event_df, caterer_df, food_df):
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.title()

# Join everything into one wide table
full_df = (
    event_orders_df
    .merge(event_df, on="event_id", how="left")
    .merge(caterer_df, on="caterer_id", how="left")
    .merge(food_df, on="food_order_id", how="left", suffixes=("_event", "_food"))
)

# Make sure numeric columns are numeric
for col in ["cost", "total_unit_cost", "quantity", "leftovers", "unit_price", "fee",
            "attendees_num", "rsvp_num"]:
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
                # Main title
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

                # High-level summary widgets
                html.Div(
                    style={
                        "display": "flex",
                        "gap": "20px",
                        "flexWrap": "wrap",
                        "justifyContent": "center",
                        "marginBottom": "30px",
                    },
                    children=[
                        html.Div(
                            style={
                                "border": "1px solid #ddd",
                                "padding": "15px",
                                "borderRadius": "10px",
                                "backgroundColor": "white",
                                "minWidth": "220px",
                                "flex": "1 1 240px",
                                "textAlign": "center",
                                "boxShadow": "0 1px 3px rgba(0,0,0,0.1)",
                            },
                            children=[
                                html.H4(
                                    "Total Events",
                                    style={"marginBottom": "10px", "fontWeight": "bold"},
                                ),
                                html.P(
                                    f"{event_metrics['event_id'].nunique()}",
                                    style={"fontSize": "20px", "fontWeight": "bold"},
                                ),
                            ],
                        ),
                        html.Div(
                            style={
                                "border": "1px solid #ddd",
                                "padding": "15px",
                                "borderRadius": "10px",
                                "backgroundColor": "white",
                                "minWidth": "220px",
                                "flex": "1 1 240px",
                                "textAlign": "center",
                                "boxShadow": "0 1px 3px rgba(0,0,0,0.1)",
                            },
                            children=[
                                html.H4(
                                    "Total Spend",
                                    style={"marginBottom": "10px", "fontWeight": "bold"},
                                ),
                                html.P(
                                    format_currency(event_metrics["total_cost"].sum()),
                                    style={
                                        "fontSize": "20px",
                                        "fontWeight": "bold",
                                        "color": ACCENT_COLOR,
                                    },
                                ),
                            ],
                        ),
                        html.Div(
                            style={
                                "border": "1px solid #ddd",
                                "padding": "15px",
                                "borderRadius": "10px",
                                "backgroundColor": "white",
                                "minWidth": "220px",
                                "flex": "1 1 240px",
                                "textAlign": "center",
                                "boxShadow": "0 1px 3px rgba(0,0,0,0.1)",
                            },
                            children=[
                                html.H4(
                                    "Average Cost Per Attendee",
                                    style={"marginBottom": "10px", "fontWeight": "bold"},
                                ),
                                html.P(
                                    format_currency(event_metrics["cost_per_attendee"].mean()),
                                    style={
                                        "fontSize": "20px",
                                        "fontWeight": "bold",
                                        "color": ACCENT_COLOR,
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
                                                            title="Cost Per Attendee vs Event Size",
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
                                                        "Total Cost": caterer_metrics["total_cost"].map(format_currency),
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
                                                            food_metrics,
                                                            x="total_quantity",
                                                            y="avg_leftovers",
                                                            hover_name="order",
                                                            title="Food Quantity Vs Average Leftovers",
                                                            labels={
                                                                "total_quantity": "Total Quantity Ordered",
                                                                "avg_leftovers": "Average Leftovers",
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
                                        "alignItems": "center",
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
                                            options=[
                                                {"label": name, "value": eid}
                                                for eid, name in event_metrics[["event_id", "event_name"]].values
                                            ],
                                            value=event_metrics["event_id"].iloc[0],
                                            style={
                                                "width": "400px",
                                                "backgroundColor": "white",
                                                "border": f"1px solid {ACCENT_COLOR}",
                                                "borderRadius": "8px",
                                            },
                                            placeholder="Search or select an event...",
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
                                            style={"width": "100%", "maxWidth": "580px", "margin": "0 auto"},
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
                    ],
                ),

                # Footer / bottom spacing
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
    event_rows = full_df[full_df["event_id"] == selected_event_id].copy()

    # Bar chart: cost by food item for this event
    food_breakdown = (
        event_rows
        .groupby("order", as_index=False)
        .agg(total_cost=("cost", "sum"), quantity=("quantity", "sum"))
    )

    fig = px.bar(
        food_breakdown,
        x="order",
        y="total_cost",
        title="Cost By Food Item For Selected Event",
        labels={"order": "Food Item", "total_cost": "Cost ($)"},
    )
    fig = style_bar(fig)
    fig.update_yaxes(tickprefix="$")

    # Detail table for this event
    display_cols = [
        "event_name", "caterer_name", "order", "dietary_label",
        "quantity", "leftovers", "unit_price", "fee", "cost",
    ]
    existing_cols = [c for c in display_cols if c in event_rows.columns]
    table_df = event_rows[existing_cols].copy()

    # Rename columns for nicer headers
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

    # Format currency columns
    for col in ["Unit Price", "Fee", "Total Cost"]:
        if col in table_df.columns:
            table_df[col] = table_df[col].map(format_currency)

    data = table_df.to_dict("records")
    columns = [{"name": col, "id": col} for col in table_df.columns]

    return fig, data, columns


# ---------- Main ----------

if __name__ == "__main__":
    app.run(debug=True)
