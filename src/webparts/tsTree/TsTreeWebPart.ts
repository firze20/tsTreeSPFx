import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {SPComponentLoader} from '@microsoft/sp-loader';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
// PNP Controls
import {IFolder, PropertyFieldFolderPicker} from '@pnp/spfx-property-controls';


import styles from './TsTreeWebPart.module.scss';

import $ from 'jquery';
import 'jstree';

//Service
import {FolderService} from '../services/folder.service';
import { IFileInfo } from '@pnp/sp/files/types';
import { IFolderInfo } from "@pnp/sp/folders";

//models

import {INode, ITreeData} from '../../models';

export interface ITsTreeWebPartProps {
  description: string;
  rootFolder: IFolder;
  selectedFolder: IFolder | undefined;
  canCreate: boolean;
  canEdit: boolean;
  canMove: boolean;
  canDelete: boolean;
  filesInfo: IFileInfo[] | undefined;
  foldersInfo: IFolderInfo[] | undefined;
  node: INode[] | undefined;
  tree: ITreeData[] | undefined;
}

export default class TsTreeWebPart extends BaseClientSideWebPart<ITsTreeWebPartProps> {

  private folderService: FolderService;


  protected async onInit(): Promise<void> {
    //Starting the folder service object as soon as the webpart gets loaded
    this.folderService = new FolderService(this.context);
    this.properties.rootFolder = await this.folderService.getRootFolder();
    //Load css js tree
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/jstree/3.2.1/themes/default/style.min.css');
    return super.onInit();
  }

  public render(): void {

    if(!this.properties.selectedFolder) {
      this.domElement.innerHTML = `
       <div class="${styles.welcome}">
        <h2>JSTree <img class="${styles.jstreeIcon}" src=${require('./assets/jstree.png')} /> </h2>
        <p>Pick a folder from the web part configuration properties.</p>
        </div>`;
    }

    else {
      this.domElement.innerHTML = `
        <div class="${styles.divTree}">
          <div id='jstree'>

          </div>
        </div>
      `;

      $('#jstree').jstree({ 'core' : {
        'data' : [
           this.properties.selectedFolder.Name,
           {
             'text' : 'Root node 2',
             'state' : {
               'opened' : true,
               'selected' : true
             },
             'children' : [
               { 'text' : 'Child 1' },
               'Child 2'
             ]
          }
        ]
    } });
    }
  
  }

  //Working
  private async setSelectedFolder(folder: IFolder): Promise<void> {
    this.properties.selectedFolder = folder;
    try {
      const filesInfo = await this.folderService.getFilesInsideFolder(this.properties.selectedFolder.ServerRelativeUrl);
      this.properties.filesInfo = filesInfo;
      const childFolderInfo = await this.folderService.getChildFolders(this.properties.selectedFolder.ServerRelativeUrl);
      this.properties.foldersInfo = childFolderInfo;
      this.mappingFoldersAndFiles();
    } catch (error) {
      console.log(error);
    }
  }

  private mappingFoldersAndFiles(): void {
    const nodeFilesAndFolders: INode[] = [];
   //mapping folder 
   this.properties.foldersInfo.forEach(folder => {
     nodeFilesAndFolders.push({
       Name: folder.Name,
       id: folder.UniqueId,
       type: 'folder'
     });
     //mapping files
     this.properties.filesInfo.forEach(file => {
       nodeFilesAndFolders.push({
         Name: file.Name,
         id: folder.UniqueId,
         type: 'file'
       });
     });

     console.log(nodeFilesAndFolders);

     this.properties.node = nodeFilesAndFolders;

     console.log(this.properties.node);

     this.mappingTree();
   });
  }

  private mappingTree(): void {
    const treeData: ITreeData[] = [];
    this.properties.node.forEach(node => {
      treeData.push({
        text: node.Name,
        state: {
          disabled: false,
          opened: true,
          selected: false
        }
      });
    });
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
            description: 'JS Tree Settings'
          },
          groups: [
            {
              //groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldFolderPicker('documents', {
                    context: this.context,
                    label: 'Select a folder',
                    onSelect: (folder) => this.setSelectedFolder(folder),
                    rootFolder: this.properties.rootFolder,
                    selectedFolder: this.properties.selectedFolder,
                    onPropertyChange: (propertyPath: string, oldValue: any, newValue: any): void  => {
                       console.log(propertyPath, oldValue, newValue);
                    },
                    properties: undefined,
                    key: 'Document'
                }),
                PropertyPaneCheckbox('Can Create?', {
                  text: 'Enable create folders?',
                  checked: this.properties.canCreate
                }),
                PropertyPaneCheckbox('Can Edit?', {
                  text: 'Enable edit folder name ?',
                  checked: this.properties.canEdit
                }),
                PropertyPaneCheckbox('Can Move?', {
                  text: 'Enable move folders?',
                  checked: this.properties.canMove
                }),
                PropertyPaneCheckbox('Can Delete?', {
                  text: 'Enable delete folders?',
                  checked: this.properties.canDelete
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
