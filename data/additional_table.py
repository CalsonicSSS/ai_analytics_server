from flask import Flask
from sqlalchemy import create_engine, text
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Database connection configuration
DB_USERNAME = os.getenv("DB_USERNAME")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_HOST = os.getenv("DB_HOST")
DB_PORT = os.getenv("DB_PORT")
DB_NAME = os.getenv("DB_NAME")

# Create database connection
db_url = f"postgresql://{DB_USERNAME}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
engine = create_engine(db_url)


def create_enriched_sor_table():
    # SQL query to create the enriched table
    create_table_query = text(
        """
    -- Drop the table if it exists
    DROP TABLE IF EXISTS "EnrichedSorDetail";
    
    -- Create the new table
    CREATE TABLE "EnrichedSorDetail" AS
    WITH unique_master AS (
        SELECT DISTINCT ON (sm."SalesOrder") 
            sm."SalesOrder",
            sm."Customer",
            sm."Salesperson",
            sm."OrderDate",
            sm."Branch",
            sm."Area",
            sm."CustomerPoNumber",
            sm."ReqShipDate",
            sm."CustomerName",
            sm."CashCredit",
            sm."Currency",
            sm."ShipAddr1",
            sm."ShipAddr2",
            sm."ShipAddr3",
            sm."ShipAddr4",
            sm."ShipAddr5",
            sm."PostalCode",
            -- Include ArCustomer fields through another join
            ac."ShortName",
            ac."MasterAccount"
        FROM "SorMasterRep" sm
        LEFT JOIN "ArCustomer" ac ON sm."Customer" = ac."Customer"
        ORDER BY sm."SalesOrder", sm."InvoiceNumber"  -- Takes the first invoice entry for each SalesOrder
    )
    SELECT 
        sd."Invoice",
        sd."SalesOrder",
        sd."SalesOrderLine",
        sd."LineType",
        sd."StockCode",
        COALESCE(im."Description", sd."StockDescription") as "StockDescription",
        sd."Warehouse",
        sd."OrderQty",
        sd."ShipQty",
        sd."BackOrderQty",
        sd."LineShipDate",
        sd."Price",
        sd."ProductClass",
        sd."BuyingGroup",
        sd."TariffCode",
        sd."Contract",
        um."Customer",
        um."Salesperson",
        um."OrderDate",
        um."Branch",
        um."Area",
        um."CustomerPoNumber",
        um."ReqShipDate",
        um."CustomerName",
        um."CashCredit",
        um."Currency",
        um."ShipAddr1",
        um."ShipAddr2",
        um."ShipAddr3",
        um."ShipAddr4",
        um."ShipAddr5",
        um."PostalCode",
        um."ShortName",
        um."MasterAccount"
    FROM "SorDetailRep" sd
    LEFT JOIN unique_master um ON sd."SalesOrder" = um."SalesOrder"
    LEFT JOIN "InvMaster" im ON sd."StockCode" = im."StockCode";
    
    -- Create indexes for better query performance
    CREATE INDEX idx_enriched_sor_detail_sales_order 
    ON "EnrichedSorDetail" ("SalesOrder");
    
    CREATE INDEX idx_enriched_sor_detail_line_ship_date 
    ON "EnrichedSorDetail" ("LineShipDate");
    
    CREATE INDEX idx_enriched_sor_detail_composite 
    ON "EnrichedSorDetail" ("Branch", "Area", "Customer", "ProductClass", "Salesperson");
    
    -- Add index for the new columns
    CREATE INDEX idx_enriched_sor_detail_customer_info
    ON "EnrichedSorDetail" ("ShortName", "MasterAccount");
    

    """
    )

    # Execute the query
    with engine.connect() as conn:
        conn.execute(create_table_query)
        conn.commit()


if __name__ == "__main__":
    try:
        create_enriched_sor_table()
        print("Successfully created EnrichedSorDetail table")
    except Exception as e:
        print(f"Error creating table: {str(e)}")
