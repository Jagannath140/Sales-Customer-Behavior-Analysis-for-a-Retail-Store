-- =====================================================================
-- 01_schema_and_load.sql
-- Creates the database, the 3 tables (with PK/FK), loads the CLEANED csv files,
-- and creates one reusable view.
-- Run from the PROJECT ROOT so the relative csv paths work:
--   mysql -u root -p --local-infile=1 < sql/01_schema_and_load.sql
-- =====================================================================
DROP DATABASE IF EXISTS retail_analytics;
CREATE DATABASE retail_analytics CHARACTER SET utf8mb4;
USE retail_analytics;

-- Parent tables first, child (Sales) last, because Sales has foreign keys to both.
CREATE TABLE Customers (
    CustomerID VARCHAR(10) PRIMARY KEY,
    Name       VARCHAR(50)  NOT NULL,
    Age        INT          NOT NULL,
    Gender     VARCHAR(10)  NOT NULL,
    City       VARCHAR(30)  NOT NULL,
    JoinDate   DATE         NOT NULL
);

CREATE TABLE Products (
    ProductID   VARCHAR(10) PRIMARY KEY,
    ProductName VARCHAR(50)   NOT NULL,
    Category    VARCHAR(30)   NOT NULL,
    UnitPrice   DECIMAL(10,2) NOT NULL,
    CostPrice   DECIMAL(10,2) NOT NULL,
    CHECK (UnitPrice > 0 AND CostPrice > 0)
);

CREATE TABLE Sales (
    SalesID    VARCHAR(10) PRIMARY KEY,
    CustomerID VARCHAR(10) NOT NULL,
    ProductID  VARCHAR(10) NOT NULL,
    Quantity   INT         NOT NULL CHECK (Quantity > 0),
    SaleDate   DATE        NOT NULL,
    FOREIGN KEY (CustomerID) REFERENCES Customers(CustomerID),
    FOREIGN KEY (ProductID)  REFERENCES Products(ProductID),
    INDEX idx_sale_date (SaleDate)   -- speeds up monthly/yearly reports
);

-- Load the cleaned files (dates are already ISO yyyy-mm-dd, so DATE columns load directly)
LOAD DATA LOCAL INFILE 'data/cleaned/Customers.csv' INTO TABLE Customers
    FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' LINES TERMINATED BY '\n' IGNORE 1 LINES;
LOAD DATA LOCAL INFILE 'data/cleaned/Products.csv' INTO TABLE Products
    FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' LINES TERMINATED BY '\n' IGNORE 1 LINES;
LOAD DATA LOCAL INFILE 'data/cleaned/Sales.csv' INTO TABLE Sales
    FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' LINES TERMINATED BY '\n' IGNORE 1 LINES;

-- Row-count check (expect 50 / 30 / 300)
SELECT 'Customers' AS tbl, COUNT(*) AS row_count FROM Customers
UNION ALL SELECT 'Products', COUNT(*) FROM Products
UNION ALL SELECT 'Sales',    COUNT(*) FROM Sales;

-- One row per sale with customer + product details and calculated Revenue / Profit.
-- Every analysis query below reuses this view so the business logic lives in ONE place.
CREATE OR REPLACE VIEW vw_sales_detail AS
SELECT s.SalesID, s.SaleDate,
       YEAR(s.SaleDate) AS SaleYear, MONTH(s.SaleDate) AS SaleMonth,
       DATE_FORMAT(s.SaleDate, '%Y-%m') AS YearMonth,
       c.CustomerID, c.Name AS CustomerName, c.Age, c.Gender, c.City, c.JoinDate,
       p.ProductID, p.ProductName, p.Category, p.UnitPrice, p.CostPrice,
       s.Quantity,
       s.Quantity * p.UnitPrice                   AS Revenue,
       s.Quantity * p.CostPrice                   AS Cost,
       s.Quantity * (p.UnitPrice - p.CostPrice)   AS Profit
FROM Sales s
JOIN Customers c ON s.CustomerID = c.CustomerID
JOIN Products  p ON s.ProductID  = p.ProductID;
