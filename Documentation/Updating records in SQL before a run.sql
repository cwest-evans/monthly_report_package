SELECT * FROM [dbo].[EGC_Pulse_table_RecipientConfig]
SELECT * FROM [dbo].[EGC_Pulse_table_RecipientScope]
SELECT * FROM [dbo].[EGC_Pulse_table_RunRecipient]
SELECT * FROM [dbo].[EGC_Pulse_table_RunRecipientFile]

EXEC dbo.EGC_sp_Pulse_Run_Init;

UPDATE [EGC_Pulse_table_RunRecipient]
SET Status = 'PENDING'
WHERE RecipientEmail IN (
('jmillard@evans-gc.com'),
('erech@evans-gc.com'),
('barslan@evans-gc.com'),
('cwest@evans-gc.com'))
AND RunId = 4;

SELECT TOP 1 *
FROM dbo.EGC_Pulse_table_Run
ORDER BY RunId DESC;

DECLARE @RunId int = (SELECT TOP 1 RunId FROM dbo.EGC_Pulse_table_Run ORDER BY RunId DESC);
EXEC dbo.EGC_sp_Pulse_Run_StageAll @RunId = @RunId;

DECLARE @RunId int = (SELECT TOP 1 RunId FROM dbo.EGC_Pulse_table_Run ORDER BY RunId DESC);

-- put ONLY these 4 in your dry-fire list
DECLARE @Testers TABLE (Email nvarchar(255));
INSERT INTO @Testers (Email) VALUES
('jmillard@evans-gc.com'),
('erech@evans-gc.com'),
('barslan@evans-gc.com'),
('cwest@evans-gc.com');

-- disable everyone else for this run (keeps their rows, but prevents processing)
UPDATE rr
SET rr.Status = 'SKIPPED',
    rr.UpdatedAt = SYSDATETIME()
FROM dbo.EGC_Pulse_table_RunRecipient rr
WHERE rr.RunId = @RunId
  AND rr.RecipientEmail NOT IN (SELECT Email FROM @Testers);

-- ensure testers are pending (so they run)
UPDATE rr
SET rr.AttemptCount = 0,
    rr.Status = 'PENDING',
    rr.ErrorMessage = NULL,
    rr.LastAttemptAt = NULL,
    rr.CompletedAt = NULL,
    rr.FilePathOrUrl = NULL,
    rr.UpdatedAt = SYSDATETIME()
FROM dbo.EGC_Pulse_table_RunRecipient rr
WHERE rr.RunId = @RunId
  AND rr.RecipientEmail IN (SELECT Email FROM @Testers);

DECLARE @RunId int = (SELECT TOP 1 RunId FROM dbo.EGC_Pulse_table_Run ORDER BY RunId DESC);

SELECT Status, COUNT(*) Cnt
FROM dbo.EGC_Pulse_table_RunRecipient
WHERE RunId = @RunId
GROUP BY Status;

DECLARE @RunId int = (SELECT TOP 1 RunId FROM dbo.EGC_Pulse_table_Run ORDER BY RunId DESC);
EXEC dbo.EGC_sp_Pulse_Run_StageAll @RunId = @RunId;

DECLARE @RunId int = (SELECT TOP 1 RunId FROM dbo.EGC_Pulse_table_Run ORDER BY RunId DESC);

UPDATE rr
SET rr.AttemptCount = 0,
    rr.Status = 'PENDING',
    rr.ErrorMessage = NULL,
    rr.LastAttemptAt = NULL,
    rr.CompletedAt = NULL,
    rr.FilePathOrUrl = NULL,
    rr.UpdatedAt = SYSDATETIME()
FROM dbo.EGC_Pulse_table_RunRecipient rr
WHERE rr.RunId = @RunId
  AND rr.RecipientEmail IN (
('jmillard@evans-gc.com'),
('erech@evans-gc.com'),
('barslan@evans-gc.com'),
('cwest@evans-gc.com')

  );