/*
    Tour_Member_M 회원별 회원번호 매칭 이력 테이블
    전제: dbo.Tour_Member_M.SeqNo가 PRIMARY KEY 또는 UNIQUE KEY여야 함
*/

IF OBJECT_ID(N'dbo.Tour_Member_Account', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Tour_Member_Account
    (
        TourMemberAccountSeqNo INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Tour_Member_Account PRIMARY KEY,

    TourMemberSeqNo INT NOT NULL,
        EventCode NVARCHAR(100) NOT NULL,
        DepartureDate DATE NULL,
        ArrivalDate DATE NULL,

        -- 매칭 당시의 원본값 보존
        MemberCodeSnapshot NVARCHAR(50) NULL,
        CustomerNameSnapshot NVARCHAR(100) NULL,

        AccountID NVARCHAR(50) NOT NULL,
        AllocationOrder INT NOT NULL,
        AllocationStatus NVARCHAR(20) NOT NULL
            CONSTRAINT DF_Tour_Member_Account_AllocationStatus DEFAULT N'ALLOCATED',
        MatchMethod NVARCHAR(30) NOT NULL
            CONSTRAINT DF_Tour_Member_Account_MatchMethod DEFAULT N'MemberCode+Name',

        AllocatedAt DATETIME2(0) NOT NULL
            CONSTRAINT DF_Tour_Member_Account_AllocatedAt DEFAULT SYSDATETIME(),
        AllocatedBy NVARCHAR(50) NULL,
        ReleasedAt DATETIME2(0) NULL,
        ReleasedBy NVARCHAR(50) NULL,
        Remark NVARCHAR(500) NULL,

        CONSTRAINT FK_Tour_Member_Account_Tour_Member_M
            FOREIGN KEY (TourMemberSeqNo)
            REFERENCES dbo.Tour_Member_M (SeqNo),

        CONSTRAINT CK_Tour_Member_Account_AllocationOrder
            CHECK (AllocationOrder > 0),

        CONSTRAINT CK_Tour_Member_Account_AllocationStatus
            CHECK (AllocationStatus IN (N'ALLOCATED', N'RELEASED', N'CANCELLED'))
    );
END;
GO

/* 동일 고객 행에 같은 순번의 회원번호 중복 저장 방지 */
IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE name = N'UX_Tour_Member_Account_Member_Order'
      AND object_id = OBJECT_ID(N'dbo.Tour_Member_Account')
)
BEGIN
    CREATE UNIQUE INDEX UX_Tour_Member_Account_Member_Order
        ON dbo.Tour_Member_Account (TourMemberSeqNo, AllocationOrder);
END;
GO

/* 같은 항차에서 동일 회원번호의 중복 배정 방지 */
IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE name = N'UX_Tour_Member_Account_Active_Account'
      AND object_id = OBJECT_ID(N'dbo.Tour_Member_Account')
)
BEGIN
    CREATE UNIQUE INDEX UX_Tour_Member_Account_Active_Account
        ON dbo.Tour_Member_Account
        (EventCode, DepartureDate, ArrivalDate, AccountID)
        WHERE AllocationStatus = N'ALLOCATED';
END;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.extended_properties
    WHERE major_id = OBJECT_ID(N'dbo.Tour_Member_Account')
      AND minor_id = 0
      AND name = N'MS_Description'
)
BEGIN
    EXEC sys.sp_addextendedproperty
        @name = N'MS_Description',
        @value = N'항차별 고객과 회원번호의 매칭 및 배정 이력',
        @level0type = N'SCHEMA', @level0name = N'dbo',
        @level1type = N'TABLE',  @level1name = N'Tour_Member_Account';
END;
GO

DECLARE @Description TABLE
(
    ColumnName SYSNAME,
    Description NVARCHAR(200)
);

INSERT INTO @Description (ColumnName, Description)
VALUES
(N'TourMemberAccountSeqNo', N'회원번호 매칭 이력 순번'),
(N'TourMemberSeqNo', N'Tour_Member_M 고객 행 순번'),
(N'EventCode', N'행사코드'),
(N'DepartureDate', N'출발일'),
(N'ArrivalDate', N'도착일'),
(N'MemberCodeSnapshot', N'매칭 당시 회원코드'),
(N'CustomerNameSnapshot', N'매칭 당시 고객명'),
(N'AccountID', N'배정된 회원번호'),
(N'AllocationOrder', N'고객별 회원번호 배정 순서'),
(N'AllocationStatus', N'배정상태(ALLOCATED/RELEASED/CANCELLED)'),
(N'MatchMethod', N'매칭 방식'),
(N'AllocatedAt', N'회원번호 배정일시'),
(N'AllocatedBy', N'회원번호 배정 작업자'),
(N'ReleasedAt', N'회원번호 해제일시'),
(N'ReleasedBy', N'회원번호 해제 작업자'),
(N'Remark', N'비고');

DECLARE
    @ColumnName SYSNAME,
    @DescriptionText NVARCHAR(200);

DECLARE DescriptionCursor CURSOR LOCAL FAST_FORWARD FOR
    SELECT ColumnName, Description
    FROM @Description;

OPEN DescriptionCursor;
FETCH NEXT FROM DescriptionCursor INTO @ColumnName, @DescriptionText;

WHILE @@FETCH_STATUS = 0
BEGIN
    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.extended_properties ep
        INNER JOIN sys.columns c
            ON ep.major_id = c.object_id
           AND ep.minor_id = c.column_id
        WHERE ep.name = N'MS_Description'
          AND c.object_id = OBJECT_ID(N'dbo.Tour_Member_Account')
          AND c.name = @ColumnName
    )
    BEGIN
        EXEC sys.sp_addextendedproperty
            @name = N'MS_Description',
            @value = @DescriptionText,
            @level0type = N'SCHEMA', @level0name = N'dbo',
            @level1type = N'TABLE',  @level1name = N'Tour_Member_Account',
            @level2type = N'COLUMN', @level2name = @ColumnName;
    END;

    FETCH NEXT FROM DescriptionCursor INTO @ColumnName, @DescriptionText;
END;

CLOSE DescriptionCursor;
DEALLOCATE DescriptionCursor;
GO
