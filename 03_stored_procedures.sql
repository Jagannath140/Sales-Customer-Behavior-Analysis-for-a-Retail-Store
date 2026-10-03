-- =====================================================================
-- 03_stored_procedures.sql
-- Reusable summary reports (required by the project document).
-- Run:  mysql -u root -p retail_analytics < sql/03_stored_procedures.sql
-- Use:  CALL sp_monthly_revenue_report(2023);
-- =====================================================================
USE retail_analytics;
DROP PROCEDURE IF EXISTS sp_monthly_revenue_report;
DROP PROCEDURE IF EXISTS sp_category_summary;
DROP PROCEDURE IF EXISTS sp_top_customers;
DROP PROCEDURE IF EXISTS sp_top_products;
DROP PROCEDURE IF EXISTS sp_customer_report;
DROP PROCEDURE IF EXISTS sp_executive_summary;

DELIMITER $$

-- 1) Monthly revenue/profit for a chosen year
CREATE PROCEDURE sp_monthly_revenue_report(IN p_year INT)
BEGIN
    SELECT YearMonth, SUM(Revenue) AS Revenue, SUM(Profit) AS Profit,
           ROUND(SUM(Profit) / SUM(Revenue) * 100, 2) AS MarginPct
    FROM vw_sales_detail
    WHERE SaleYear = p_year
    GROUP BY YearMonth
    ORDER BY YearMonth;
END$$

-- 2) Category performance (no input)
CREATE PROCEDURE sp_category_summary()
BEGIN
    SELECT Category, SUM(Quantity) AS Units, SUM(Revenue) AS Revenue, SUM(Profit) AS Profit,
           ROUND(SUM(Profit) / SUM(Revenue) * 100, 2) AS MarginPct
    FROM vw_sales_detail
    GROUP BY Category
    ORDER BY Revenue DESC;
END$$

-- 3) Top N customers by spend
CREATE PROCEDURE sp_top_customers(IN p_n INT)
BEGIN
    SELECT CustomerID, CustomerName, City, COUNT(*) AS Orders,
           SUM(Revenue) AS TotalSpent, SUM(Profit) AS ProfitGenerated
    FROM vw_sales_detail
    GROUP BY CustomerID, CustomerName, City
    ORDER BY TotalSpent DESC
    LIMIT p_n;
END$$

-- 4) Top N products by revenue, optionally filtered by category (pass NULL for all)
CREATE PROCEDURE sp_top_products(IN p_n INT, IN p_category VARCHAR(30))
BEGIN
    SELECT ProductName, Category, SUM(Quantity) AS Units, SUM(Revenue) AS Revenue, SUM(Profit) AS Profit
    FROM vw_sales_detail
    WHERE p_category IS NULL OR Category = p_category
    GROUP BY ProductID, ProductName, Category
    ORDER BY Revenue DESC
    LIMIT p_n;
END$$

-- 5) Everything about one customer: profile line + their purchases
CREATE PROCEDURE sp_customer_report(IN p_customer_id VARCHAR(10))
BEGIN
    SELECT CustomerID, CustomerName, City, COUNT(*) AS Orders, SUM(Quantity) AS Units,
           SUM(Revenue) AS CLV_Revenue, SUM(Profit) AS CLV_Profit,
           MIN(SaleDate) AS FirstPurchase, MAX(SaleDate) AS LastPurchase
    FROM vw_sales_detail WHERE CustomerID = p_customer_id
    GROUP BY CustomerID, CustomerName, City;

    SELECT SaleDate, ProductName, Category, Quantity, Revenue, Profit
    FROM vw_sales_detail WHERE CustomerID = p_customer_id
    ORDER BY SaleDate;
END$$

-- 6) One-call executive summary: KPIs, best month, best product, best city
CREATE PROCEDURE sp_executive_summary()
BEGIN
    SELECT SUM(Revenue) AS TotalRevenue, SUM(Profit) AS TotalProfit,
           ROUND(SUM(Profit) / SUM(Revenue) * 100, 2) AS ProfitMarginPct,
           SUM(Quantity) AS TotalUnits, COUNT(DISTINCT CustomerID) AS ActiveCustomers
    FROM vw_sales_detail;

    SELECT YearMonth AS BestMonth, SUM(Revenue) AS Revenue
    FROM vw_sales_detail GROUP BY YearMonth ORDER BY Revenue DESC LIMIT 1;

    SELECT ProductName AS BestProductByRevenue, SUM(Revenue) AS Revenue
    FROM vw_sales_detail GROUP BY ProductID, ProductName ORDER BY Revenue DESC LIMIT 1;

    SELECT City AS BestCity, SUM(Revenue) AS Revenue
    FROM vw_sales_detail GROUP BY City ORDER BY Revenue DESC LIMIT 1;
END$$

DELIMITER ;
