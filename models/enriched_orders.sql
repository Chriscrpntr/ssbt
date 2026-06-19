SELECT
    o.order_id,
    o.customer,
    o.region,
    o.total,
    c.email
FROM {{ ref('completed_orders') }} o
JOIN {{ ref('customers_customer_list') }} c ON o.customer = c.name
WHERE o.status = 'completed'