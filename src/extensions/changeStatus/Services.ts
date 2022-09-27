import { IItemAddResult, IItems, IItemUpdateResult, ISiteGroupInfo, IWeb, sp, Web } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import { IDocStatus, IListItem } from "./Types";
import { FieldCustomizerContext } from "@microsoft/sp-listview-extensibility";

export class listService {

    public async getListItem(listTitle: string, id: number): Promise<IListItem> {
        const select: string = 'FileRef,FileLeafRef,*';
        const listItem: IItems = await sp.web.lists.getByTitle(listTitle).items.getById(id).select(select)();
        return convertType('IListItem', listItem);
    }

    public async saveListItem(listTitle: string, id: number, fields:any): Promise<IItemUpdateResult> {
        const result:IItemUpdateResult = await sp.web.lists.getByTitle(listTitle).items.getById(id).update(fields);
        return result;
    }

    public async isAdmin(adminGroup: string): Promise<boolean> {
        const groups: ISiteGroupInfo[] = await sp.web.currentUser.groups();
        let result: boolean = false;
        groups.forEach(group => {
            if (group.Title === adminGroup) {
                result = true;
            }
        });
        return result;
    }

    public async getDocStatusList(context: FieldCustomizerContext): Promise<IDocStatus[]> {
        const root: IWeb = Web(context.pageContext.site.absoluteUrl);
        const docStatusList: any[] = await root.getList(context.pageContext.site.serverRelativeUrl + '/Lists/StatusConfig').items();
        return docStatusList;
    }

    public async saveWorkflowTechnialListItem(context: FieldCustomizerContext, fields: any): Promise<IItemAddResult> {
        const root: IWeb = Web(context.pageContext.site.absoluteUrl);
        const newTechItem:IItemAddResult = await root.getList(context.pageContext.site.serverRelativeUrl + '/Lists/WorkflowTechnicalList').items.add(fields);
        return newTechItem;
    }
}

const SPListViewService = new listService();
export default SPListViewService;

function convertType(type: string, listItem: any): IListItem | PromiseLike<IListItem> {
    switch (type) {
        case 'IListItem':
            return {
                Id: Number(listItem.ID),
                Name: String(listItem.FileLeafRef),
                LinkToItem: String(listItem.FileRef),
                PeriodId: listItem.PeriodId ? listItem.PeriodId : null,
                PeriodText: listItem.PeriodText ? String(listItem.PeriodText) : null,
                DocStatusId: listItem.DocStatusId ? listItem.DocStatusId : null,
                DocStatusText: listItem.DocStatusText ? String(listItem.DocStatusText) : null,
                Comment: listItem.Comment1 ? String(listItem.Comment1) : null,
                PermissionSetInProgress: listItem.PermissionSetInProgress ? Number(listItem.PermissionSetInProgress) : 0
            };
        default:
            return listItem;
    }
}
