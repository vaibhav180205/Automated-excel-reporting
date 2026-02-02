#!/usr/bin/env python3
"""
Database Setup Script

This script creates a SQLite database with sample sales data.
Run this script first before running the main reporting script.

Usage: python setup_database.py
"""

import sqlite3
import random
from datetime import datetime, timedelta

def create_database():
    """
    Create SQLite database and populate with sample sales data.
    """
    print("Creating SQLite database...")
    
    # Connect to database (creates file if doesn't exist)
    conn = sqlite3.connect('sales_data.db')
    cursor = conn.cursor()
    
    # Create sales table
    print("Creating sales table...")
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS sales (
        sale_id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_date TEXT NOT NULL,
        product_name TEXT NOT NULL,
        category TEXT NOT NULL,
        quantity INTEGER NOT NULL,
        unit_price REAL NOT NULL,
        total_amount REAL NOT NULL
    )
    ''')
    
    # Sample products and categories
    products = [
        ('Laptop', 'Electronics', 899.99),
        ('Smartphone', 'Electronics', 599.99),
        ('Headphones', 'Electronics', 79.99),
        ('Office Chair', 'Furniture', 249.99),
        ('Desk', 'Furniture', 399.99),
        ('Monitor', 'Electronics', 299.99),
        ('Keyboard', 'Electronics', 49.99),
        ('Mouse', 'Electronics', 29.99),
        ('Bookshelf', 'Furniture', 179.99),
        ('Table Lamp', 'Furniture', 39.99)
    ]
    
    # Generate sample sales data for the past 90 days
    print("Generating sample sales data...")
    sales_data = []
    start_date = datetime.now() - timedelta(days=90)
    
    for day in range(90):
        current_date = start_date + timedelta(days=day)
        date_str = current_date.strftime('%Y-%m-%d')
        
        # Generate random number of sales per day (5-15)
        num_sales = random.randint(5, 15)
        
        for _ in range(num_sales):
            # Pick random product
            product_name, category, unit_price = random.choice(products)
            
            # Random quantity (1-5)
            quantity = random.randint(1, 5)
            
            # Calculate total
            total_amount = quantity * unit_price
            
            sales_data.append((
                date_str,
                product_name,
                category,
                quantity,
                unit_price,
                total_amount
            ))
    
    # Insert data into table
    cursor.executemany('''
    INSERT INTO sales (sale_date, product_name, category, quantity, unit_price, total_amount)
    VALUES (?, ?, ?, ?, ?, ?)
    ''', sales_data)
    
    # Commit changes and close connection
    conn.commit()
    
    # Display summary
    cursor.execute('SELECT COUNT(*) FROM sales')
    total_records = cursor.fetchone()[0]
    
    cursor.execute('SELECT SUM(total_amount) FROM sales')
    total_revenue = cursor.fetchone()[0]
    
    print(f"\n{'='*60}")
    print("DATABASE CREATED SUCCESSFULLY!")
    print(f"{'='*60}")
    print(f"Total sales records: {total_records}")
    print(f"Total revenue: ${total_revenue:,.2f}")
    print(f"Database file: sales_data.db")
    print(f"{'='*60}\n")
    
    conn.close()

if __name__ == "__main__":
    create_database()
