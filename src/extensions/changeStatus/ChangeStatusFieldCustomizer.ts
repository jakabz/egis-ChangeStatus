import * as React from 'react';
import * as ReactDOM from 'react-dom';

import {
  BaseFieldCustomizer,
  FieldCustomizerContext,
  IFieldCustomizerCellEventParameters,
  ListItemAccessor
} from '@microsoft/sp-listview-extensibility';

import ChangeStatus, { IChangeStatusProps } from './components/ChangeStatus';

export interface IChangeStatusFieldCustomizerProperties {
  AdminGroup: string;
}

export default class ChangeStatusFieldCustomizer
  extends BaseFieldCustomizer<IChangeStatusFieldCustomizerProperties> {

  public onInit(): Promise<void> {
    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const row: ListItemAccessor = event.listItem;
    const context: FieldCustomizerContext = this.context;
    const adminGroup: string = this.properties.AdminGroup;
    const changeStatus: React.ReactElement<{}> =
      React.createElement(ChangeStatus, { row, context, adminGroup } as IChangeStatusProps);

    ReactDOM.render(changeStatus, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
