export interface Result {
    PrimaryQueryResult: PrimaryQueryResult;
}

export interface PrimaryQueryResult{
    RelevantResults: RelevantResults;
}

export interface RelevantResults {
    Table: Table;
    TotalRows: number;
}

export interface Table {
    Rows: Row[];
}

export interface Row {
    Cells: Cell[];
    length: number;
}

export interface Cell {
    Key: string;
    Value: string;
}