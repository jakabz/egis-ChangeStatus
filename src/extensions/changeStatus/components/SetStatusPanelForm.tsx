import { DefaultButton, Dropdown, IDropdownOption, Panel, PrimaryButton, Stack, Text, TextField } from 'office-ui-fabric-react';
import * as React from 'react';
import { FieldCustomizerContext, ListItemAccessor } from '@microsoft/sp-listview-extensibility';
import SPListViewService from '../Services';
import { IDocStatus, IListItem } from '../Types';

export interface ISetStatusPanelFormProps {
    row: ListItemAccessor;
    closePanel: (openPanel: boolean) => void;
    context: FieldCustomizerContext;
    adminGroup: string;
}

export const SetStatusPanelForm: React.FunctionComponent<ISetStatusPanelFormProps> = (props: React.PropsWithChildren<ISetStatusPanelFormProps>) => {

    const [listItem, setListItem] = React.useState<IListItem>(null);
    const [isAdmin, setIsAdmin] = React.useState<boolean>(false);
    const [docStatusList, setDocStatusList] = React.useState<IDocStatus[]>(null);
    const [statusField, setStatusField] = React.useState<string | number>(null);
    const [commentField, setCommentField] = React.useState<string>(null);

    React.useEffect(() => {
        SPListViewService.getListItem(props.context.pageContext.list.title, props.row.getValueByName('ID'))
            .then(item => {
                setListItem(item);
                setStatusField(item.DocStatusId);
                setCommentField(item.Comment);
            })
            .catch(err => console.info(err));
        SPListViewService.isAdmin(props.adminGroup).then(result => setIsAdmin(result)).catch(err => console.info(err));
        SPListViewService.getDocStatusList(props.context).then(result => setDocStatusList(result)).catch(err => console.info(err));
    }, []);

    const saveForm = (): void => {
        SPListViewService.saveListItem(props.context.pageContext.list.title, listItem.Id, {
            Comment1: commentField,
            PermissionSetInProgress: listItem.DocStatusId === statusField ? 0 : 1
        }).then(resultListItem => {
            if (listItem.DocStatusId !== statusField) {
                SPListViewService.saveWorkflowTechnialListItem(props.context, {
                    Title: "Modification",
                    FileName: listItem.Name,
                    LinkToItem: location.origin + listItem.LinkToItem,
                    ItemID: listItem.Id,
                    DocumentStatusID: statusField,
                    PeriodID: listItem.PeriodId
                }).then(resultTechListItem => {
                    props.closePanel(false);
                }).catch(err => console.log(err));
            } else {
                props.closePanel(false);
            }
        }).catch(err => console.info(err));
    };


    const onRenderFooterContent = (): JSX.Element => {
        return (
            <Stack horizontal gap={5} style={{ padding: 20 }}>
                <PrimaryButton onClick={saveForm} disabled={listItem?.PermissionSetInProgress === 1}>
                    Save
                </PrimaryButton>
                <DefaultButton onClick={() => props.closePanel(false)}>Cancel</DefaultButton>
            </Stack>
        );
    };

    const docStatusOptions = (): IDropdownOption[] => {
        const options: IDropdownOption[] = [];
        if (isAdmin) {
            docStatusList?.forEach((status) => {
                options.push({ key: status.Id, text: status.Title });
            });
        } else {
            const currDocStatus = docStatusList?.find((status) => status.Id === listItem.DocStatusId);
            if (currDocStatus) {
                options.push({ key: currDocStatus.Id, text: currDocStatus.Title });
            }
            currDocStatus?.FollowingStatusesId.forEach(nextStatusId => {
                const nextStatusItem = docStatusList?.find((status) => status.Id === nextStatusId);
                options.push({ key: nextStatusItem.Id, text: nextStatusItem.Title });
            });
        }
        return options;
    }

    return (
        <Panel
            headerText="Set status"
            isOpen={true}
            onDismiss={() => props.closePanel(false)}
            closeButtonAriaLabel="Close"
            onRenderFooter={onRenderFooterContent}
            isFooterAtBottom={true}
        >
            {
                listItem && listItem.PermissionSetInProgress === 0 ? (
                    <Stack>
                        <TextField
                            label="Name"
                            disabled={true}
                            value={listItem.Name}
                        />
                        <TextField
                            label="Period"
                            disabled={true}
                            value={listItem.PeriodText}
                        />
                        <Dropdown
                            placeholder="Select options"
                            label="Document status"
                            selectedKey={statusField}
                            required={false}
                            options={docStatusOptions()}
                            onChange={(event, option, index) => setStatusField(option.key)}
                        />
                        <TextField
                            label="Comment"
                            multiline={true}
                            rows={5}
                            disabled={false}
                            required={true}
                            value={commentField}
                            onChange={(event, newValue) => setCommentField(newValue)}
                        />
                    </Stack>
                )
                    :
                    (
                        <Stack>
                            <Text>A dokumentumon jogosultság állítása van folyamatban, kérjük próbálja meg később szerkeszteni.</Text>
                        </Stack>
                    )
            }
        </Panel>
    );
};