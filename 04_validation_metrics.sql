-- =====================================================================
-- 04_validation_metrics.sql
-- Key control totals used to cross-check MySQL against Excel, Python and Power BI.
-- Run:  mysql -u root -B retail_analytics < sql/04_validation_metrics.sql > outputs/mysql_validation_metrics.tsv
-- =====================================================================
USE retail_analytics;
SELECT 'Total Revenue' AS Metric, SUM(Revenue) AS Value FROM vw_sales_detail
UNION ALL SELECT 'Total Profit', SUM(Profit) FROM vw_sales_detail
UNION ALL SELECT 'Total Units', SUM(Quantity) FROM vw_sales_detail
UNION ALL SELECT 'Transactions', COUNT(*) FROM vw_sales_detail
UNION ALL SELECT 'Active Customers', COUNT(DISTINCT CustomerID) FROM vw_sales_detail
UNION ALL SELECT 'Products Sold', COUNT(DISTINCT ProductID) FROM vw_sales_detail
UNION ALL SELECT 'Electronics Revenue', SUM(Revenue) FROM vw_sales_detail WHERE Category = 'Electronics'
UNION ALL SELECT 'Accessories Revenue', SUM(Revenue) FROM vw_sales_detail WHERE Category = 'Accessories'
UNION ALL SELECT 'Dubai Revenue', SUM(Revenue) FROM vw_sales_detail WHERE City = 'Dubai'
UNION ALL SELECT 'Sep-2023 Revenue', SUM(Revenue) FROM vw_sales_detail WHERE YearMonth = '2023-09'
UNION ALL SELECT 'Customer C046 Revenue', SUM(Revenue) FROM vw_sales_detail WHERE CustomerID = 'C046'
UNION ALL SELECT 'Camera DSLR Revenue', SUM(Revenue) FROM vw_sales_detail WHERE ProductName = 'Camera DSLR';
