SELECT a.id, a.type, SUM(t.amount) AS total
FROM accounts a
JOIN transactions t ON t.account_id = a.id
WHERE t.posted_at >= DATE '2026-01-01'
GROUP BY a.id, a.type
ORDER BY total DESC;
