/*
Page 1: Sessions (SessionAdmin)

The PostgreSQL driver supplies the start date and end date as two parameters.
The end date is inclusive because the filter stops at the next calendar day.
*/
WITH filtered_sessions AS (
    SELECT
        id,
        project_code,
        username,
        created_at,
        closed_at
        FROM documents_session
        WHERE created_at >= %s::date
            AND created_at < %s::date + INTERVAL '1 day'
            AND username NOT IN ('ron.lee.ks@asiaalliedgroup.com','ken.chui@hypath.ai','cyrus.lo@asiaalliedgroup.com')
),
event_stats AS (
    SELECT
        session_id,
        COUNT(*) FILTER (WHERE action = 'RETRIEVE') AS count_retrieve_mode,
        COUNT(DISTINCT (meta->>'method_statement_id')) FILTER (WHERE action = 'PDF') AS count_pdf_viewed,
        COUNT(*) FILTER (WHERE action = 'COPY') AS count_content_copied,
        COUNT(*) FILTER (WHERE action = 'ASK') AS count_ask_mode,
        COUNT(*) FILTER (
            WHERE action = 'SUBMIT'
              AND meta->>'file' LIKE '%%.docx%%'
        ) AS count_submission_times
    FROM documents_events
    WHERE session_id IN (SELECT id FROM filtered_sessions)
    GROUP BY session_id
),
report_stats AS (
    SELECT
        session_id,
        COUNT(*) FILTER (WHERE status IN (0, 1)) AS pending_count,
        COUNT(*) AS total_count
    FROM documents_report
    WHERE session_id IN (SELECT id FROM filtered_sessions)
    GROUP BY session_id
)
SELECT
    s.project_code,
    s.username,
    s.created_at,
    s.closed_at,
    COALESCE(es.count_retrieve_mode, 0) AS count_retrieve_mode,
    COALESCE(es.count_pdf_viewed, 0) AS count_pdf_viewed,
    COALESCE(es.count_content_copied, 0) AS count_content_copied,
    COALESCE(es.count_ask_mode, 0) AS count_ask_mode,
    COALESCE(es.count_submission_times, 0) AS count_submission_times,
    COALESCE(rs.total_count, 0) AS total_issue_count,
    CONCAT(COALESCE(rs.pending_count, 0), ' / ', COALESCE(rs.total_count, 0)) AS pending_issue
FROM filtered_sessions s
LEFT JOIN event_stats es ON s.id = es.session_id
LEFT JOIN report_stats rs ON s.id = rs.session_id
ORDER BY s.created_at DESC;
