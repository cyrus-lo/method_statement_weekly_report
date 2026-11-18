/*
Page 2: Submission Counts (per Unique Submission ID)

The PostgreSQL driver supplies the start date and end date as two parameters.
The end date is inclusive because the filter stops at the next calendar day.
*/
SELECT
    s.project_code,
    COUNT(DISTINCT (e.meta->>'submission_id')) AS count_submission_times
FROM documents_session s
JOIN documents_events e ON s.id = e.session_id
WHERE e.action = 'SUBMIT'
  AND e.meta->>'file' LIKE '%%.docx%%'
  AND s.created_at >= %s::date
  AND s.created_at < %s::date + INTERVAL '1 day'
GROUP BY s.project_code
HAVING COUNT(DISTINCT (e.meta->>'submission_id')) > 0
ORDER BY s.project_code;
