import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

// PNP Controls

import {IFolder, PropertyFieldFolderPicker} from '@pnp/spfx-property-controls';

import styles from './TsTreeWebPart.module.scss';
import * as strings from 'TsTreeWebPartStrings';

//Service

import {FolderService} from '../services/folder.service';

export interface ITsTreeWebPartProps {
  description: string;
  context: WebPartContext;
  rootFolder: IFolder;
  selectedFolder: IFolder;
}

export default class TsTreeWebPart extends BaseClientSideWebPart<ITsTreeWebPartProps> {

  private folderService: FolderService;


  protected async onInit(): Promise<void> {
    //Starting the folder service object as soon as the webpart gets loaded
    this.folderService = new FolderService(this.context);
    this.properties.rootFolder = await this.folderService.getRootFolder();
    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.welcome}">
       <h2>JSTree</h2>
       <p>Pick a folder from the web part configuration properties.</p>
      </div>`;
  }

  //Working
  private setSelectedFolder(folder: IFolder) {
    console.log(folder.Name);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldFolderPicker('documents', {
                    context: this.context,
                    label: 'Select a folder',
                    onSelect: this.setSelectedFolder,
                    rootFolder: this.properties.rootFolder,
                    selectedFolder: undefined,
                    onPropertyChange: function (propertyPath: string, oldValue: any, newValue: any): void {
                       console.log(propertyPath, oldValue, newValue);
                    },
                    properties: undefined,
                    key: 'Document'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
