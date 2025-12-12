DROP TABLE IF EXISTS Scans;
CREATE TABLE Scans (
    ID INTEGER PRIMARY KEY,
    ArtifactID INTEGER NOT NULL CHECK(ArtifactID > 0),
    ScanTool TEXT NOT NULL CHECK(length(trim(ScanTool)) > 0),
    ScanType TEXT NOT NULL CHECK(length(trim(ScanType)) > 0),
    ScanDateTime TEXT NOT NULL CHECK(length(trim(ScanDateTime)) > 0),
    ScanRepeatCount INTEGER NOT NULL DEFAULT 1,
    Critical INTEGER NOT NULL DEFAULT 0,
    High INTEGER NOT NULL DEFAULT 0,
    Medium INTEGER NOT NULL DEFAULT 0,
    CriticalNP INTEGER NOT NULL DEFAULT 0,
    HighNP INTEGER NOT NULL DEFAULT 0,
    MediumNP INTEGER NOT NULL DEFAULT 0,
    CONSTRAINT UniqueVTDate UNIQUE(ArtifactID, ScanTool, ScanType, ScanDateTime),
    FOREIGN KEY (ArtifactID) REFERENCES Artifacts(ID)
);
