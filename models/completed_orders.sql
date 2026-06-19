SELECT * FROM {{ ref('raw_orders') }} WHERE status = 'completed'
