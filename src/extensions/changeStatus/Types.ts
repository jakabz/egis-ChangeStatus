export interface IListItem {
    Id: number;
    Name: string;
    LinkToItem: string;
    PeriodId: number;
    PeriodText: string;
    DocStatusId: number;
    DocStatusText: string;
    Comment: string;
    PermissionSetInProgress: number;
}

export interface IDocStatus {
    Id: number;
    Title: string;
    FollowingStatusesId: number[];
}