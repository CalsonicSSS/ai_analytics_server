import pandas as pd

# SQLAlchemy is a SQL toolkit and Object-Relational Mapping (ORM) library for Python.
# It is widely used for interacting, managing with relational databases (like PostgreSQL, MySQL, SQLite, etc.) within Python applications in a more flexible and Pythonic way
# SQLAlchemy provides a high-level, Pythonic interface to interact with PostgreSQL databases, making it easier to write database operations in Python.
from sqlalchemy import create_engine
from dotenv import load_dotenv
import os
import load_data_type_config

# Load environment variables
load_dotenv()

# PostgreSQL connection string
conn_string = f"postgresql://{os.getenv('DB_USERNAME')}:{os.getenv('DB_PASSWORD')}@{os.getenv('DB_HOST')}/{os.getenv('DB_NAME')}"

# Create SQLAlchemy engine
engine = create_engine(conn_string)

# Get list of CSV files in the data folder
csv_files = [f for f in os.listdir("csv") if f.endswith(".csv")]


# Load each CSV file into PostgreSQL
for file in csv_files:
    print(f"Loading {file}...")
    file_path = os.path.join("csv", file)
    table_name = file.split(".")[0]

    dtypes = load_data_type_config.custom_dtypes.get(table_name, {})

    # Read CSV with custom dtypes
    df = pd.read_csv(file_path, dtype=dtypes)

    df.to_sql(table_name, engine, if_exists="replace", index=False)
    print(f"Loaded {file} into table {table_name}")

print("Data loading completed successfully")
