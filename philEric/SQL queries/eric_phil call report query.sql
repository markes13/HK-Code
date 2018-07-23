USE CCSuiteDB
;

SELECT
	Agent,
	COUNT(DISTINCT "Caller ID") As "Unique Incoming Calls",
	COUNT("Type") AS "Incoming Calls",
	AVG("Call Duration")/60 AS "Avg Call Duration (Minutes)",
	SUM("Call Duration")/60 AS "Sum of Call Duration (Minutes)",
	AVG("Talk Time")/60 AS "Avg Talk Time (Minutes)",
	AVG("Ring Time") AS "Avg Ring Time (Seconds)",
	Type
FROM
	dbo.rCOCallLog
WHERE
	Agent IN ('Eric Thomas', 'Phil Trietsch')
	-- get past week
	AND CAST(TIME AS DATE) > CAST(GETDATE() - 8 AS DATE)
	AND CAST(TIME AS DATE) < CAST(GETDATE() AS DATE)
	AND "Type" IN (
	'ACD', 'Non ACD', 'Lost', 'Out'
	)
GROUP BY Agent, "Type"
ORDER BY Agent
;