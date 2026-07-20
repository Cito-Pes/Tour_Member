IF COL_LENGTH(N'dbo.Tour_Member_M', N'TravelRegion') IS NULL
BEGIN
    ALTER TABLE dbo.Tour_Member_M
        ADD TravelRegion NVARCHAR(100) NULL;
END;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.extended_properties ep
    INNER JOIN sys.columns c
        ON ep.major_id = c.object_id
       AND ep.minor_id = c.column_id
    WHERE ep.name = N'MS_Description'
      AND c.object_id = OBJECT_ID(N'dbo.Tour_Member_M')
      AND c.name = N'TravelRegion'
)
BEGIN
    EXEC sys.sp_addextendedproperty
        @name = N'MS_Description',
        @value = N'여행지역',
        @level0type = N'SCHEMA', @level0name = N'dbo',
        @level1type = N'TABLE',  @level1name = N'Tour_Member_M',
        @level2type = N'COLUMN', @level2name = N'TravelRegion';
END;
GO
