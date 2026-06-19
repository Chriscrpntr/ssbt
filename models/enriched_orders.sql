SELECT
    o.order_id,
    o.customer,
    o.region,
    o.status,
    o.total,
    c.email,
    c.tier
FROM completed_orders o
JOIN {{ ref('customers_customer_list') }} c ON o.customer = c.name
