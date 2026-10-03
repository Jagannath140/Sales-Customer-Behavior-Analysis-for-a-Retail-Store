-- =====================================================================
-- 02_analysis_queries.sql
-- Business questions answered with SQL. Each query states the question it answers.
-- Run:  mysql -u root -p retail_analytics < sql/02_analysis_queries.sql
-- =====================================================================
USE retail_analytics;

-- Q1. What are our headline numbers? (KPI cards)
SELECT SUM(Revenue)                                   AS TotalRevenue,
       SUM(Profit)                                    AS TotalProfit,
       ROUND(SUM(Profit) / SUM(Revenue) * 100, 2)     AS ProfitMarginPct,
       SUM(Quantity)                                  AS TotalUnits,
       COUNT(*)                                       AS TotalTransactions,
       COUNT(DISTINCT CustomerID)                     AS ActiveCustomers,
       COUNT(DISTINCT ProductID)                      AS ProductsSold
FROM vw_sales_detail;

-- Q2. How does revenue move month to month? (seasonality / peak months)
--     LAG() gives the previous month so we can compute month-on-month growth.
SELECT YearMonth,
       SUM(Revenue) AS Revenue,
       SUM(Profit)  AS Profit,
       ROUND(SUM(Profit) / SUM(Revenue) * 100, 2) AS MarginPct,
       ROUND((SUM(Revenue) - LAG(SUM(Revenue)) OVER (ORDER BY YearMonth))
             / LAG(SUM(Revenue)) OVER (ORDER BY YearMonth) * 100, 1) AS MoM_Growth_Pct
FROM vw_sales_detail
GROUP BY YearMonth
ORDER BY YearMonth;

-- Q3. Which products sell the most? (best sellers by units, with revenue for context)
SELECT ProductName, Category,
       SUM(Quantity) AS UnitsSold,
       SUM(Revenue)  AS Revenue,
       SUM(Profit)   AS Profit
FROM vw_sales_detail
GROUP BY ProductID, ProductName, Category
ORDER BY UnitsSold DESC, Revenue DESC
LIMIT 10;

-- Q4. Which products make the most money? (by revenue) and at what margin?
SELECT ProductName, Category,
       SUM(Revenue) AS Revenue,
       SUM(Profit)  AS Profit,
       ROUND(SUM(Profit) / SUM(Revenue) * 100, 2) AS MarginPct,
       RANK() OVER (ORDER BY SUM(Revenue) DESC)   AS RevenueRank
FROM vw_sales_detail
GROUP BY ProductID, ProductName, Category
ORDER BY Revenue DESC;

-- Q5. Which category drives revenue vs profit margin?
SELECT Category,
       SUM(Quantity) AS Units,
       SUM(Revenue)  AS Revenue,
       SUM(Profit)   AS Profit,
       ROUND(SUM(Profit) / SUM(Revenue) * 100, 2)                   AS MarginPct,
       ROUND(SUM(Revenue) / (SELECT SUM(Revenue) FROM vw_sales_detail) * 100, 1) AS RevenueSharePct
FROM vw_sales_detail
GROUP BY Category
ORDER BY Revenue DESC;

-- Q6. Who are the high-value customers? (top 10 by total spend)
SELECT CustomerID, CustomerName, City,
       COUNT(*)      AS Orders,
       SUM(Revenue)  AS TotalSpent,
       SUM(Profit)   AS ProfitGenerated
FROM vw_sales_detail
GROUP BY CustomerID, CustomerName, City
ORDER BY TotalSpent DESC
LIMIT 10;

-- Q7. Customer segments: NTILE(5) splits customers into 5 equal groups by spend; group 1 = top 20 %.
SELECT CASE WHEN spend_quintile = 1 THEN 'High value (top 20%)'
            WHEN spend_quintile = 5 THEN 'Low value (bottom 20%)'
            ELSE 'Mid value' END AS Segment,
       COUNT(*)        AS Customers,
       SUM(TotalSpent) AS Revenue,
       ROUND(SUM(TotalSpent) / SUM(SUM(TotalSpent)) OVER () * 100, 1) AS RevenueSharePct
FROM (SELECT CustomerID, SUM(Revenue) AS TotalSpent,
             NTILE(5) OVER (ORDER BY SUM(Revenue) DESC) AS spend_quintile
      FROM vw_sales_detail GROUP BY CustomerID) t
GROUP BY 1
ORDER BY Revenue DESC;

-- Q8. Customer Lifetime Value (historical): total profit each customer generated in the data window.
--     With only 12 months and no churn data, a predictive CLV would need invented assumptions.
SELECT CustomerID, CustomerName,
       COUNT(*)                          AS Orders,
       SUM(Revenue)                      AS CLV_Revenue,
       SUM(Profit)                       AS CLV_Profit,
       ROUND(SUM(Revenue) / COUNT(*), 2) AS AvgOrderValue,
       MIN(SaleDate) AS FirstPurchase, MAX(SaleDate) AS LastPurchase
FROM vw_sales_detail
GROUP BY CustomerID, CustomerName
ORDER BY CLV_Profit DESC;

-- Q9. Which cities generate the most revenue, and how much per customer?
SELECT City,
       COUNT(DISTINCT CustomerID)                                AS Customers,
       SUM(Revenue)                                              AS Revenue,
       SUM(Profit)                                               AS Profit,
       ROUND(SUM(Revenue) / COUNT(DISTINCT CustomerID), 0)       AS RevenuePerCustomer
FROM vw_sales_detail
GROUP BY City
ORDER BY Revenue DESC;

-- Q10. Gender split of revenue (see data note: name/gender pairs look inconsistent, treat as indicative)
SELECT Gender, COUNT(DISTINCT CustomerID) AS Customers, SUM(Revenue) AS Revenue, SUM(Profit) AS Profit
FROM vw_sales_detail GROUP BY Gender;

-- Q11. Which products have thin margins? (document suggests flagging < 10 %; we also show < 20 %)
SELECT ProductName, Category, UnitPrice, CostPrice,
       ROUND((UnitPrice - CostPrice) / UnitPrice * 100, 2) AS MarginPct,
       CASE WHEN (UnitPrice - CostPrice) / UnitPrice < 0.10 THEN 'Below 10%'
            WHEN (UnitPrice - CostPrice) / UnitPrice < 0.20 THEN 'Below 20%'
            ELSE 'OK' END AS MarginFlag
FROM Products
ORDER BY MarginPct;

-- Q12. Purchasing behaviour: how often, how much, how many distinct products does each customer buy?
SELECT CustomerID, CustomerName,
       COUNT(*) AS Orders, SUM(Quantity) AS Units,
       COUNT(DISTINCT ProductID)  AS DistinctProducts,
       COUNT(DISTINCT YearMonth)  AS ActiveMonths
FROM vw_sales_detail
GROUP BY CustomerID, CustomerName
ORDER BY Orders DESC
LIMIT 10;

-- Q13. Data-quality check: sales dated BEFORE the customer's join date (flagged in Python too)
SELECT COUNT(*) AS SalesBeforeJoinDate FROM vw_sales_detail WHERE SaleDate < JoinDate;
