SELECT
    region,
    COUNT(*) AS order_count,
    SUM(total) AS total_revenue,
    AVG(total) AS avg_order_value
FROM {{ ref('enriched_orders') }}
GROUP BY region
