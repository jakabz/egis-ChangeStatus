import { FieldCustomizerContext, ListItemAccessor } from '@microsoft/sp-listview-extensibility';
import { ActionButton } from 'office-ui-fabric-react';
import * as React from 'react';

import styles from './ChangeStatus.module.scss';
import { SetStatusPanelForm } from './SetStatusPanelForm';

export interface IChangeStatusProps {
  row: ListItemAccessor;
  context: FieldCustomizerContext;
  adminGroup: string;
}

export interface IChangeStatusState {
  openPanel: boolean;
}

export default class ChangeStatus extends React.Component<IChangeStatusProps, IChangeStatusState> {

  constructor(props: IChangeStatusProps) {
    super(props);
    this.state = {
      openPanel: false
    };
  }

  public render(): React.ReactElement<{}> {

    const openStatusPanel = (): void => {
      this.setState({ openPanel: true });
    };

    return (
      <React.Fragment>
        <ActionButton
          iconProps={{ iconName: 'NavigateForward' }}
          allowDisabledFocus
          className={styles.SetStatusButton}
          onClick={openStatusPanel}
        >
          Set status
        </ActionButton>
        {
          this.state.openPanel && <SetStatusPanelForm
              row={this.props.row}
              context={this.props.context}
              adminGroup={this.props.adminGroup}
              closePanel={(openPanel) => this.setState({ openPanel: openPanel })}
            />
        }

      </React.Fragment>
    );
  }
}