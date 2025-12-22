import os
from dotenv import load_dotenv
import psycopg2

load_dotenv()

database = os.getenv("POSTGRES_DB")
user = os.getenv("POSTGRES_USER")
password = os.getenv("POSTGRES_PASSWORD")

def db_conection():
    conn = psycopg2.connect(
    host="localhost",
    database=database,
    user=user,
    password=password,
    port=5432)
    return conn
