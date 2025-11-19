import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import ListDirectory from './components/ListDirectory';
import { IListDirectoryProps } from './components/IListDirectoryProps';

export default class ListDirectoryWebPart extends BaseClientSideWebPart<IListDirectoryProps> {
  
  public render(): void {
    const element: React.ReactElement<IListDirectoryProps> = React.createElement(
      ListDirectory,
      {
        context: this.context
      }
    );

    ReactDOM.render(element, this.domElement);
  }

  // Unmount the React component to avoid memory leaks
  public onDispose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
