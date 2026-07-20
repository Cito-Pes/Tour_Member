DECLARE @Description TABLE
(
    ColumnName SYSNAME NOT NULL,
    DescriptionText NVARCHAR(200) NOT NULL
);

INSERT INTO @Description (ColumnName, DescriptionText)
VALUES
(N'PrePay', N'선납금'),
(N'SupPay', N'선지원금액'),
(N'PortCh', N'항만세'),
(N'FuelSurCha', N'유류할증료'),
(N'ExcRateAddCha', N'환율추가금'),
(N'IntraTip', N'선내팁'),
(N'RoomUpYN', N'객실업그레이드 유무'),
(N'RoomUpCha', N'객실업그레이드 비용'),
(N'AirUpYN', N'항공 업그레이드 유무'),
(N'AirUpCha', N'항공업그레이드 비용'),
(N'VisaCha', N'비자비용'),
(N'RoomDCCha', N'객실할인비용'),
(N'AirAddCha', N'항공료추가비용'),
(N'SpeAddCha', N'스페셜추가비용'),
(N'AirDCCha', N'항공할인비용'),
(N'OneTimePay', N'일시납 금액'),
(N'Penalty', N'위약금'),
(N'EtcCha1', N'기타비용');

DECLARE
    @ColumnName SYSNAME,
    @DescriptionText NVARCHAR(200),
    @ObjectId INT = OBJECT_ID(N'dbo.Event_Expenses');

IF @ObjectId IS NULL
BEGIN
    RAISERROR(N'dbo.Event_Expenses 테이블이 없습니다.', 16, 1);
    RETURN;
END;

DECLARE DescriptionCursor CURSOR LOCAL FAST_FORWARD FOR
    SELECT ColumnName, DescriptionText
    FROM @Description;

OPEN DescriptionCursor;
FETCH NEXT FROM DescriptionCursor INTO @ColumnName, @DescriptionText;

WHILE @@FETCH_STATUS = 0
BEGIN
    IF COL_LENGTH(N'dbo.Event_Expenses', @ColumnName) IS NOT NULL
    BEGIN
        IF EXISTS
        (
            SELECT 1
            FROM sys.extended_properties ep
            INNER JOIN sys.columns c
                ON ep.major_id = c.object_id
               AND ep.minor_id = c.column_id
            WHERE ep.name = N'MS_Description'
              AND c.object_id = @ObjectId
              AND c.name = @ColumnName
        )
        BEGIN
            EXEC sys.sp_updateextendedproperty
                @name = N'MS_Description',
                @value = @DescriptionText,
                @level0type = N'SCHEMA', @level0name = N'dbo',
                @level1type = N'TABLE',  @level1name = N'Event_Expenses',
                @level2type = N'COLUMN', @level2name = @ColumnName;
        END
        ELSE
        BEGIN
            EXEC sys.sp_addextendedproperty
                @name = N'MS_Description',
                @value = @DescriptionText,
                @level0type = N'SCHEMA', @level0name = N'dbo',
                @level1type = N'TABLE',  @level1name = N'Event_Expenses',
                @level2type = N'COLUMN', @level2name = @ColumnName;
        END;
    END;

    FETCH NEXT FROM DescriptionCursor INTO @ColumnName, @DescriptionText;
END;

CLOSE DescriptionCursor;
DEALLOCATE DescriptionCursor;
