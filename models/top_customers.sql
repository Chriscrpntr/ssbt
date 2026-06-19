SELECT
    customer,
    COUNT(*) AS order_count,
    SUM(total) AS total_spent
FROM {{ ref('enriched_orders') }}
GROUP BY customer
ORDER BY total_spent DESC
LIMIT 10
