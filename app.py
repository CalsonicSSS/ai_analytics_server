from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
from flask_caching import Cache
from flask import send_file
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
import calendar
from openpyxl.styles import Alignment, Font, PatternFill
from urllib.parse import quote_plus


# SQLAlchemy is a SQL toolkit and Object-Relational Mapping (ORM) library for Python.
# It is widely used as high-level, Pythonic interface for interacting, managing with relational databases (like PostgreSQL, MySQL, SQLite, etc.) within Python applications in a more flexible and Pythonic way
from sqlalchemy import create_engine, text
from dotenv import load_dotenv
import os
import math

# Load environment variables
# this function reads the .env file (default in root level) and loads those variables into the current environment (not permanently added to your system environment).
# The environment variables from the .env file only affect the current Python process and are not permanent system-wide variables.
load_dotenv()

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# PostgreSQL connection string
# conn_string_local = f"postgresql://{os.getenv('DB_USERNAME')}:{os.getenv('DB_PASSWORD')}@{os.getenv('DB_HOST')}/{os.getenv('DB_NAME')}"
conn_string_railway = f"postgresql://{os.getenv('DB_USERNAME')}:{os.getenv('DB_PASSWORD')}@{os.getenv('DB_HOST')}:{os.getenv('DB_PORT')}/{os.getenv('DB_NAME')}?sslmode=require"

# Create SQLAlchemy engine: It handles the details of connecting to and communicating with the database. It serves as the starting point for any database operations in SQLAlchemy.
engine = create_engine(conn_string_railway)

# Configure Flask-Caching
# Flask-Caching is a simple caching extension for Flask to add caching support for various backends to any Flask application.
cache = Cache(app, config={"CACHE_TYPE": "SimpleCache", "CACHE_DEFAULT_TIMEOUT": 86400})  # 1 day in seconds

# -------------------------------------------------------------------------------------------------------------------------------------------------------------


@app.route("/", methods=["GET"])
def home():
    print("/ route reached")
    # using jsonify to convert the Python dictionary to a JSON string as a response payload and send directly to the client side after return statement
    # below is the structure of customizing the response headers, status code and response payload
    return jsonify({"message": "Welcome to flowmatic"}), 200, {"Custom-Header1": "custom value 1"}


# decorator in Flask that maps the URL path '/api/order_status_overview' to the function that follows it (When a request is made to this URL, Flask will execute the function below it)
# below function is called a view function, contain the logic to be executed when the corresponding URL is accessed. It returns a response to the client.
@app.route("/api/order_status_overview")
def order_status_overview():
    query = text(
        """
        WITH status_counts AS (
            SELECT 
                "OrderStatus",
                COUNT(*) as Count,
                COUNT(*) * 100.0 / SUM(COUNT(*)) OVER () as Percentage
            FROM 
                "SorMasterRep"
            GROUP BY 
                "OrderStatus"
        ),
        all_statuses(status) AS (
            VALUES ('0'), ('1'), ('2'), ('3'), ('4'), ('8'), ('9'), ('F')
        )
        SELECT 
            CASE 
                WHEN a.status = '0' THEN 'Order in process'
                WHEN a.status = '1' THEN 'Open order'
                WHEN a.status = '2' THEN 'Open back order'
                WHEN a.status = '3' THEN 'Released back order'
                WHEN a.status = '4' THEN 'In warehouse'
                WHEN a.status = '8' THEN 'To invoice'
                WHEN a.status = '9' THEN 'Complete'
                WHEN a.status = 'F' THEN 'Forwarded order'
            END as "statusDesc",
            a.status as "statusCode",
            COALESCE(s.Count, 0) as "count",
            ROUND(COALESCE(s.Percentage, 0), 2) as "percentage"
        FROM 
            all_statuses a
        LEFT JOIN 
            status_counts s ON a.status = s."OrderStatus"
        ORDER BY 
            "count" DESC;
        """
    )

    with engine.connect() as conn:
        # result is a SQLAlchemy CursorResult: You can loop over it to get each row of the result. It has methods like keys() to get column names and fetchall() to get all rows at once.
        # Each row in the result is a Row object, which is similar to a tuple but also allows dictionary-like access.
        result = conn.execute(query)
        columns = result.keys()
        data = [dict(zip(columns, row)) for row in result]

    return jsonify({"status": "success", "data": data})


@app.route("/api/order_qty_filter_options")
@cache.cached(timeout=86400)  # Cache this endpoint for 1 hour
def get_filter_options():
    """Endpoint to fetch all available filter options"""
    with engine.connect() as conn:
        filter_query = text(
            """
            SELECT 
                ARRAY_AGG(DISTINCT "Branch") AS branches,
                ARRAY_AGG(DISTINCT "Area") AS areas,
                ARRAY_AGG(DISTINCT "Customer") AS customers,
                ARRAY_AGG(DISTINCT "ProductClass") AS product_classes,
                ARRAY_AGG(DISTINCT "Salesperson") AS salespersons
            FROM "EnrichedSorDetail"
        """
        )
        result = conn.execute(filter_query).fetchone()

    return jsonify(
        {
            "status": "success",
            "data": {
                "branch": sorted(filter(None, result.branches)),
                "area": sorted(filter(None, result.areas)),
                "customer": sorted(filter(None, result.customers)),
                "product_class": sorted(filter(None, result.product_classes)),
                "salesperson": sorted(filter(None, result.salespersons)),
            },
        }
    )


@app.route("/api/order_qty_trend_data")
def order_qty_trend():

    # Get filter parameters (all optional)
    filters = {
        "branch": request.args.get("branch"),
        "area": request.args.get("area"),
        "customer": request.args.get("customer"),
        "product_class": request.args.get("product_class"),
        "salesperson": request.args.get("salesperson"),
    }

    with engine.connect() as conn:
        # Query to get trend data with full table date range
        trend_query = text(
            """
            WITH full_date_range AS (
                -- Get min and max dates from ENTIRE table (before filtering)
                SELECT 
                    DATE_TRUNC('month', MIN(TO_DATE("OrderDate", 'MM/DD/YYYY'))) AS min_date,
                    DATE_TRUNC('month', MAX(TO_DATE("OrderDate", 'MM/DD/YYYY'))) AS max_date
                FROM "EnrichedSorDetail"
            ),
            date_series AS (
                -- Generate series of months for the entire date range
                SELECT generate_series(
                    (SELECT min_date FROM full_date_range),
                    (SELECT max_date FROM full_date_range),
                    '1 month'::interval
                ) AS month
            ),
            monthly_totals AS (
                -- Calculate monthly totals with filters
                SELECT 
                    DATE_TRUNC('month', TO_DATE("OrderDate", 'MM/DD/YYYY')) AS order_month,
                    SUM("OrderQty") AS total_qty
                FROM "EnrichedSorDetail"
                WHERE 1=1
                    AND (:branch IS NULL OR "Branch" = :branch)
                    AND (:area IS NULL OR "Area" = :area)
                    AND (:customer IS NULL OR "Customer" = :customer)
                    AND (:product_class IS NULL OR "ProductClass" = :product_class)
                    AND (:salesperson IS NULL OR "Salesperson" = :salesperson)
                GROUP BY DATE_TRUNC('month', TO_DATE("OrderDate", 'MM/DD/YYYY'))
            )
            SELECT 
                TO_CHAR(ds.month, 'YYYY-MM') AS "yearMonth",
                COALESCE(mt.total_qty, 0) AS "totalOrderQty"
            FROM date_series ds
            LEFT JOIN monthly_totals mt ON ds.month = mt.order_month
            ORDER BY ds.month ASC
        """
        )

        # Execute query with filters
        trend_data = [dict(row._mapping) for row in conn.execute(trend_query, filters)]

    return jsonify({"status": "success", "data": trend_data})


@app.route("/api/order_detail_table_data")
def order_detail():
    # Get filter parameters
    filters = {
        "branch": request.args.get("branch"),
        "area": request.args.get("area"),
        "customer": request.args.get("customer"),
        "product_class": request.args.get("product_class"),
        "salesperson": request.args.get("salesperson"),
        "start_date": request.args.get("start_date"),
        "end_date": request.args.get("end_date"),
        "page": int(request.args.get("page", 1)),
        "per_page": int(request.args.get("per_page", 50)),
    }

    with engine.connect() as conn:
        # Validate dates if provided
        if not filters["start_date"] or not filters["end_date"]:
            date_range = conn.execute(
                text(
                    """
                SELECT 
                    DATE_TRUNC('month', MIN(TO_DATE("OrderDate", 'MM/DD/YYYY'))) AS min_date,
                    (DATE_TRUNC('month', MAX(TO_DATE("OrderDate", 'MM/DD/YYYY'))) + INTERVAL '1 month' - INTERVAL '1 day')::date AS max_date
                FROM "EnrichedSorDetail"
            """
                )
            ).fetchone()
            filters["start_date"] = date_range.min_date.strftime("%m/%d/%Y")
            filters["end_date"] = date_range.max_date.strftime("%m/%d/%Y")

        # Calculate offset
        filters["offset"] = (filters["page"] - 1) * filters["per_page"]

        # Get paginated detail data
        detail_query = text(
            """
            SELECT 
                "Invoice" AS "invoice",
                "Customer" AS "customerID",
                "CustomerName" AS "customerName", 
                "ShortName" AS "shortName",
                "MasterAccount" AS "masterAccount",
                "SalesOrder" AS "sorID",
                "SalesOrderLine" AS "sorLine",
                "LineType" AS "lineType",
                "StockCode" AS "stockCode",
                "StockDescription" AS "stockDesc",
                "OrderQty" AS "orderQty",
                "ShipQty" AS "shipQty",
                "BackOrderQty" AS "backOrderQty",
                "ProductClass" AS "productClass",
                "Branch" AS "branch",
                "Area" AS "area",
                CONCAT_WS(', ', "ShipAddr1", "ShipAddr2", "ShipAddr3", "ShipAddr4", "ShipAddr5", "PostalCode") AS "shippingAddress",
                "OrderDate" AS "orderDate",
                "Salesperson" AS "salespersonID"
            FROM "EnrichedSorDetail"
            WHERE TO_DATE("OrderDate", 'MM/DD/YYYY') >= TO_DATE(:start_date, 'MM/DD/YYYY')
            AND TO_DATE("OrderDate", 'MM/DD/YYYY') <= TO_DATE(:end_date, 'MM/DD/YYYY')
            AND (:branch IS NULL OR "Branch" = :branch)
            AND (:area IS NULL OR "Area" = :area)
            AND (:customer IS NULL OR "Customer" = :customer)
            AND (:product_class IS NULL OR "ProductClass" = :product_class)
            AND (:salesperson IS NULL OR "Salesperson" = :salesperson)
            ORDER BY TO_DATE("OrderDate", 'MM/DD/YYYY') DESC, "SalesOrder", "SalesOrderLine"
            LIMIT :per_page OFFSET :offset
        """
        )

        detail_data = [dict(row._mapping) for row in conn.execute(detail_query, filters)]

        # Get total count for pagination
        count_query = text(
            """
            SELECT COUNT(*)
            FROM "EnrichedSorDetail"
            WHERE TO_DATE("OrderDate", 'MM/DD/YYYY') >= TO_DATE(:start_date, 'MM/DD/YYYY')
            AND TO_DATE("OrderDate", 'MM/DD/YYYY') <= TO_DATE(:end_date, 'MM/DD/YYYY')
            AND (:branch IS NULL OR "Branch" = :branch)
            AND (:area IS NULL OR "Area" = :area)
            AND (:customer IS NULL OR "Customer" = :customer)
            AND (:product_class IS NULL OR "ProductClass" = :product_class)
            AND (:salesperson IS NULL OR "Salesperson" = :salesperson)
        """
        )

        total_count = conn.execute(count_query, filters).scalar()

    return jsonify(
        {
            "status": "success",
            "data": detail_data,
            "pagination": {
                "totalCount": total_count,
                "page": filters["page"],
                "perPage": filters["per_page"],
                "totalPages": math.ceil(total_count / filters["per_page"]),
            },
        }
    )


@app.route("/api/top_order_stock_items")
def top_order_stock_items():
    query = text(
        """
            WITH TopStockItems AS (
                SELECT 
                    sd."StockCode" as "stockCode",
                    sd."StockDescription" as "stockDescription",
                    SUM(sd."OrderQty") as "totalOrderQty"
                FROM 
                    "SorDetailRep" sd
                GROUP BY 
                    sd."StockCode", sd."StockDescription"
                ORDER BY 
                    SUM(sd."OrderQty") DESC
                LIMIT 30
            )
            SELECT 
                t."stockCode",
                COALESCE(t."stockDescription", im."Description") as "description",
                t."totalOrderQty",
                im."Supplier" as "supplier",
                im."ProductClass" as "productClass",
                STRING_AGG(DISTINCT iw."Warehouse" || ': ' || iw."QtyOnHand"::text, ', ') as "warehouseQty",
                STRING_AGG(DISTINCT ip."PriceCode" || ': $' || ip."SellingPrice"::text, ', ') as "priceDetails",
                SUM(iw."QtyOnHand") as "totalQtyOnHand",
                (
                    SELECT JSON_AGG(
                        json_build_object(
                            'priceCode', ph."PriceCode",
                            'dateChanged', ph."DateChanged",
                            'timeChanged', ph."TimeChanged",
                            'newPrice', ph."NewSellingPrice",
                            'oldPrice', ph."OldSellingPrice"
                        ) ORDER BY ph."DateChanged" DESC, ph."TimeChanged" DESC
                    )
                    FROM "InvPriceHistory" ph
                    WHERE ph."StockCode" = t."stockCode"
                ) as "priceHistory"
            FROM 
                TopStockItems t
                LEFT JOIN "InvMaster" im ON t."stockCode" = im."StockCode"
                LEFT JOIN "InvWarehouse" iw ON t."stockCode" = iw."StockCode"
                LEFT JOIN "InvPrice" ip ON t."stockCode" = ip."StockCode"
            GROUP BY 
                t."stockCode",
                t."stockDescription",
                t."totalOrderQty",
                im."Description",
                im."Supplier",
                im."ProductClass"
            ORDER BY 
                t."totalOrderQty" DESC
        """
    )

    with engine.connect() as conn:
        result = conn.execute(query)
        data = [dict(row._mapping) for row in result]

        # Format the response for frontend consumption
        formatted_data = [
            {
                "stockCode": row["stockCode"],
                "description": row["description"],
                "totalOrderQty": float(row["totalOrderQty"]),
                "supplier": row["supplier"],
                "productClass": row["productClass"],
                "warehouseDetails": row["warehouseQty"],
                "priceDetails": row["priceDetails"],
                "totalQtyOnHand": float(row["totalQtyOnHand"]) if row["totalQtyOnHand"] else 0,
                "priceHistory": row["priceHistory"] if row["priceHistory"] else [],
            }
            for row in data
        ]

    return jsonify({"status": "success", "data": formatted_data})


@app.route("/api/order_by_product_class")
def order_by_product_class():
    query_top5 = text(
        """
        WITH ProductClassTotals AS (
            SELECT 
                sd."ProductClass" as "productClass",
                SUM(sd."OrderQty") as "totalOrderQty",
                RANK() OVER (ORDER BY SUM(sd."OrderQty") DESC) as rank
            FROM 
                "SorDetailRep" sd
            WHERE 
                sd."ProductClass" IS NOT NULL
            GROUP BY 
                sd."ProductClass"
        )
        SELECT 
            "productClass",
            "totalOrderQty",
            CASE 
                WHEN rank <= 5 THEN rank
                ELSE 6
            END as rank_category
        FROM 
            ProductClassTotals
        """
    )

    with engine.connect() as conn:
        result_top5 = conn.execute(query_top5)
        data_top5 = [dict(row._mapping) for row in result_top5]

        top_5 = []
        rest_value = 0
        total_order_qty = sum(row["totalOrderQty"] for row in data_top5)

        for row in data_top5:
            if row["rank_category"] <= 5:
                top_5.append(row)
            else:
                rest_value += row["totalOrderQty"]

        if rest_value > 0:
            top_5.append({"productClass": "Others", "totalOrderQty": rest_value, "rank_category": 6})

    top_5_product_classes = [row["productClass"] for row in top_5 if row["productClass"] != "Others"]

    query_descriptions = text(
        """
        SELECT 
            sp."ProductClass" as "productClass",
            COALESCE(sp."Description", 'Unknown') as "description"
        FROM 
            "SalProductClass" sp
        WHERE 
            sp."ProductClass" IN :top_5_classes
        """
    )

    with engine.connect() as conn:
        result_descriptions = conn.execute(query_descriptions, {"top_5_classes": tuple(top_5_product_classes)})

        # Fetch rows as tuples and access by index
        descriptions = {row[0]: row[1] for row in result_descriptions.fetchall()}

    formatted_data = []
    for row in top_5:
        if row["productClass"] == "Others":
            label = "Others"
        else:
            label = f"{descriptions.get(row['productClass'], 'Unknown')} ({row['productClass']})"

        percentage = round((row["totalOrderQty"] / total_order_qty) * 100, 2)

        formatted_data.append({"id": row["productClass"], "label": label, "value": round(float(row["totalOrderQty"])), "percentage": percentage})

    total_percentage = round(sum(item["percentage"] for item in formatted_data), 2)

    return jsonify({"status": "success", "data": formatted_data, "totalPercentage": total_percentage})


@app.route("/api/download_comprehensive_report")
def download_comprehensive_report():
    try:
        # Get current date information dynamically
        current_date = datetime.now()
        current_year = current_date.year
        current_month = current_date.month

        # Create the base query template for all categories
        base_query_template = """
        WITH DateRanges AS (
            -- All-time total
            SELECT 
                '{category}' as category_name,
                {column} as category_value,
                SUM("OrderQty") as total_all_time,
                
                -- Current month this year
                SUM(CASE 
                    WHEN EXTRACT(YEAR FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = {current_year}
                    AND EXTRACT(MONTH FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = {current_month}
                    THEN "OrderQty" ELSE 0 END) as current_month_qty,
                
                -- Same month last year
                SUM(CASE 
                    WHEN EXTRACT(YEAR FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = {current_year} - 1
                    AND EXTRACT(MONTH FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = {current_month}
                    THEN "OrderQty" ELSE 0 END) as last_year_month_qty,
                
                -- Same month two years ago
                SUM(CASE 
                    WHEN EXTRACT(YEAR FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = {current_year} - 2
                    AND EXTRACT(MONTH FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = {current_month}
                    THEN "OrderQty" ELSE 0 END) as two_years_ago_month_qty,
                
                -- Past 6 months detail
                SUM(CASE 
                    WHEN TO_DATE("OrderDate", 'MM/DD/YYYY') >= DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '5 months'
                    AND EXTRACT(MONTH FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = {current_month}
                    THEN "OrderQty" ELSE 0 END) as month_0_qty,
                
                SUM(CASE 
                    WHEN TO_DATE("OrderDate", 'MM/DD/YYYY') >= DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '5 months'
                    AND EXTRACT(MONTH FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = 
                        CASE 
                            WHEN {current_month} - 1 < 1 THEN 12
                            ELSE {current_month} - 1
                        END
                    THEN "OrderQty" ELSE 0 END) as month_1_qty,
                
                SUM(CASE 
                    WHEN TO_DATE("OrderDate", 'MM/DD/YYYY') >= DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '5 months'
                    AND EXTRACT(MONTH FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = 
                        CASE 
                            WHEN {current_month} - 2 < 1 THEN 12 + ({current_month} - 2)
                            ELSE {current_month} - 2
                        END
                    THEN "OrderQty" ELSE 0 END) as month_2_qty,
                
                SUM(CASE 
                    WHEN TO_DATE("OrderDate", 'MM/DD/YYYY') >= DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '5 months'
                    AND EXTRACT(MONTH FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = 
                        CASE 
                            WHEN {current_month} - 3 < 1 THEN 12 + ({current_month} - 3)
                            ELSE {current_month} - 3
                        END
                    THEN "OrderQty" ELSE 0 END) as month_3_qty,
                
                SUM(CASE 
                    WHEN TO_DATE("OrderDate", 'MM/DD/YYYY') >= DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '5 months'
                    AND EXTRACT(MONTH FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = 
                        CASE 
                            WHEN {current_month} - 4 < 1 THEN 12 + ({current_month} - 4)
                            ELSE {current_month} - 4
                        END
                    THEN "OrderQty" ELSE 0 END) as month_4_qty,
                
                SUM(CASE 
                    WHEN TO_DATE("OrderDate", 'MM/DD/YYYY') >= DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '5 months'
                    AND EXTRACT(MONTH FROM TO_DATE("OrderDate", 'MM/DD/YYYY')) = 
                        CASE 
                            WHEN {current_month} - 5 < 1 THEN 12 + ({current_month} - 5)
                            ELSE {current_month} - 5
                        END
                    THEN "OrderQty" ELSE 0 END) as month_5_qty
                
            FROM "EnrichedSorDetail"
            WHERE {column} IS NOT NULL 
              AND {column} != ''  -- Exclude empty strings
              AND TRIM({column}) != ''  -- Exclude whitespace-only strings
            GROUP BY {column}
            ORDER BY SUM("OrderQty") DESC
        )
        SELECT 
            category_name,
            category_value,
            total_all_time,
            current_month_qty,
            last_year_month_qty,
            two_years_ago_month_qty,
            month_0_qty,
            month_1_qty,
            month_2_qty,
            month_3_qty,
            month_4_qty,
            month_5_qty
        FROM DateRanges
        """

        # Define categories and their corresponding columns
        categories = [
            ("Branch", "Branch"),
            ("Area", "Area"),
            ("Customer", "Customer"),
            ("StockCode", "StockCode"),
            ("ProductClass", "ProductClass"),
            ("Salesperson", "Salesperson"),
        ]

        # Get month names for the past 6 months
        def get_month_name(month_number):
            # Handle month number wrapping
            if month_number < 1:
                month_number = 12 + month_number
            return calendar.month_name[((month_number - 1) % 12) + 1]

        current_month_name = get_month_name(current_month)

        # Calculate past 6 months considering year wrap-around
        month_names = []
        for i in range(6):
            month_num = current_month - i
            if month_num < 1:
                month_num = 12 + month_num
            month_names.append(get_month_name(month_num))

        # Create Excel writer object
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Process each category
            for category_name, column_name in categories:
                # Execute query for current category
                query = base_query_template.format(
                    category=category_name, column=f'"{column_name}"', current_year=current_year, current_month=current_month
                )

                with engine.connect() as conn:
                    df = pd.read_sql_query(query, conn)

                # Rename columns for clarity
                df.columns = [
                    "Category",
                    f"{category_name}",
                    "Total Quantity",
                    f"{current_month_name} {current_year}",
                    f"{current_month_name} {current_year-1}",
                    f"{current_month_name} {current_year-2}",
                ] + [f"{month} {current_year if current_month-i >= 1 else current_year-1}" for i, month in enumerate(month_names)]

                # Sort by total quantity descending
                df = df.sort_values("Total Quantity", ascending=False)

                # Write to Excel sheet
                df.to_excel(writer, sheet_name=category_name, index=False)

                # Get the worksheet
                worksheet = writer.sheets[category_name]

                # Define maximum column widths
                max_widths = {
                    "Category": 15,  # Category name
                    f"{category_name}": 30,  # Category value (allow slightly longer for codes/names)
                    "Total Quantity": 15,  # Numeric columns
                    "Month columns": 15,  # All month-based columns
                }

                # Apply formatting to each column
                for idx, col in enumerate(df.columns):
                    # Get column letter
                    col_letter = chr(65 + idx)

                    # Calculate width (with minimum and maximum limits)
                    if idx == 1:  # Category value column
                        width = min(max(len(str(col)), df[col].astype(str).apply(len).max(), 8), max_widths[f"{category_name}"])  # minimum width
                    else:  # Other columns
                        width = min(max(len(str(col)), df[col].astype(str).apply(len).max(), 8), max_widths.get("Month columns", 15))  # minimum width

                    # Set column width
                    worksheet.column_dimensions[col_letter].width = width + 2

                    # Format header
                    header_cell = worksheet[f"{col_letter}1"]
                    header_cell.font = Font(bold=True)
                    header_cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
                    header_cell.alignment = Alignment(wrap_text=True)

                    # Format data cells
                    for row in range(2, len(df) + 2):
                        cell = worksheet[f"{col_letter}{row}"]

                        # Align numeric columns to right
                        if idx >= 2:  # Quantity columns
                            cell.alignment = Alignment(horizontal="right")

                            # Format numbers with thousands separator and no decimal places
                            if isinstance(cell.value, (int, float)):
                                cell.number_format = "#,##0"
                        else:
                            cell.alignment = Alignment(horizontal="left")

                # Freeze the header row
                worksheet.freeze_panes = "A2"

                # Auto-filter for all columns
                worksheet.auto_filter.ref = worksheet.dimensions

        # Prepare the file for download
        output.seek(0)

        # Generate filename with current timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"comprehensive_report_{timestamp}.xlsx"

        return send_file(
            output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=filename
        )
    except Exception as e:
        print(f"Error generating report: {str(e)}")
        return jsonify({"status": "error", "message": "Failed to generate report. Please try again or contact support."}), 500


import os
from flask import send_from_directory, jsonify
from datetime import datetime


@app.route("/api/download_static_report")
def download_comprehensive_order_report():
    try:
        # Define the reports directory path
        reports_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reports")

        # Ensure the reports directory exists
        if not os.path.exists(reports_dir):
            return jsonify({"status": "error", "message": "Reports directory not found"}), 404

        # Get the latest report file from the reports directory
        report_files = [f for f in os.listdir(reports_dir) if f.endswith(".xlsx")]

        if not report_files:
            return jsonify({"status": "error", "message": "No report files available"}), 404

        # Get the most recent file based on file creation time
        latest_report = max(report_files, key=lambda f: os.path.getctime(os.path.join(reports_dir, f)))

        # Get the file creation time for display
        file_ctime = datetime.fromtimestamp(os.path.getctime(os.path.join(reports_dir, latest_report)))
        file_ctime_str = file_ctime.strftime("%Y-%m-%d %H:%M:%S")

        response = send_from_directory(
            reports_dir, latest_report, as_attachment=True, download_name=f"comprehensive_order_report_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )

        # Add cache control headers
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"

        return response

    except Exception as e:
        print(f"Error serving static report: {str(e)}")
        return jsonify({"status": "error", "message": "Failed to download report. Please try again or contact support."}), 500


if __name__ == "__main__":
    app.run(debug=True)
