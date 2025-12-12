DROP TABLE IF EXISTS Artifacts;
CREATE TABLE Artifacts (
    ID INTEGER PRIMARY KEY,
    BusinessUnit TEXT NOT NULL CHECK(length(trim(BusinessUnit)) > 0),
    AlteraProduct TEXT CHECK(AlteraProduct IS NULL OR length(trim(AlteraProduct)) > 0),
    Rapid7App TEXT CHECK(Rapid7App IS NULL OR length(trim(Rapid7App)) > 0),
    CheckmarxProduct TEXT CHECK(CheckmarxProduct IS NULL OR length(trim(CheckmarxProduct)) > 0),
    MendProduct TEXT CHECK(MendProduct IS NULL OR length(trim(MendProduct)) > 0),
    MendProject TEXT CHECK(MendProject IS NULL OR length(trim(MendProject)) > 0),
    Owner TEXT CHECK(Owner IS NULL OR length(trim(Owner)) > 0),
    SCAScans INTEGER DEFAULT 0,
    SASTScans INTEGER DEFAULT 0,
    DASTScans INTEGER DEFAULT 0,
    RecentSCA TEXT CHECK(RecentSCA IS NULL OR length(trim(RecentSCA)) > 0),
    RecentSCAOK INTEGER CHECK(RecentSCAOK IN (0, 1)) DEFAULT 0,
    RecentSAST TEXT CHECK(RecentSAST IS NULL OR length(trim(RecentSAST)) > 0),
    RecentSASTOK INTEGER CHECK(RecentSASTOK IN (0, 1)) DEFAULT 0,
    RecentDAST TEXT CHECK(RecentDAST IS NULL OR length(trim(RecentDAST)) > 0),
    RecentDASTOK INTEGER CHECK(RecentDASTOK IN (0, 1)) DEFAULT 0,
    RecentLOC INTEGER DEFAULT 0,
    Deleted INTEGER CHECK(Deleted IN (0, 1)) DEFAULT 0,

    CONSTRAINT UniqueBUR7 UNIQUE(BusinessUnit, Rapid7App),
    CONSTRAINT UniqueBUCmark UNIQUE(BusinessUnit, CheckmarxProduct),
    CONSTRAINT UniqueMprodMproj UNIQUE(MendProduct, MendProject),

    -- Require at least one data source present
    -- Rapid7 and Checkmarx: single non-empty field
    -- Mend: both product AND project must be non-empty
    CONSTRAINT OneSourcePresent CHECK (
        (Rapid7App IS NOT NULL AND length(trim(Rapid7App)) > 0)
        OR (CheckmarxProduct IS NOT NULL AND length(trim(CheckmarxProduct)) > 0)
        OR ((MendProduct IS NOT NULL AND length(trim(MendProduct)) > 0)
            AND (MendProject IS NOT NULL AND length(trim(MendProject)) > 0))
    )
);

-- Indexes
CREATE INDEX IF NOT EXISTS idx_Artifacts_BusinessUnit ON Artifacts(BusinessUnit);
