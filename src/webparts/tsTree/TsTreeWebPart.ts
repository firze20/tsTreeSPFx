import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneToggle} from '@microsoft/sp-property-pane';
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
  expandAll: boolean;
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

  //old folder and new folder

  private oldFolder: IFolder | undefined;
  private newFolder: IFolder | undefined;

  protected async onInit(): Promise<void> {
    //Starting the folder service object as soon as the webpart gets loaded
    this.folderService = new FolderService(this.context);
    this.properties.rootFolder = await this.folderService.getRootFolder();
    //Load css js tree
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/jstree/3.2.1/themes/default/style.min.css');
    return super.onInit();
  }

  public render(): void {
    
    /*var arrayCollection = [
      {"id": "animal", "parent": "#", "text": "Animals"},
      {"id": "device", "parent": "#", "text": "Devices"},
      {"id": "dog", "parent": "animal", "text": "Dogs"},
      {"id": "lion", "parent": "animal", "text": "Lions"},
      {"id": "mobile", "parent": "device", "text": "Mobile Phones"},
      {"id": "lappy", "parent": "device", "text": "Laptops"},
      {"id": "daburman", "parent": "dog", "text": "Dabur Man", "icon": "/"},
      {"id": "Dalmation", "parent": "dog", "text": "Dalmatian", "icon": "/"},
      {"id": "african", "parent": "lion", "text": "African Lion", "icon": "/"},
      {"id": "indian", "parent": "lion", "text": "Indian Lion", "icon": "/"},
      {"id": "apple", "parent": "mobile", "text": "Apple IPhone 6", "icon": "/"},
      {"id": "samsung", "parent": "mobile", "text": "Samsung Note II", "icon": "/"},
      {"id": "lenevo", "parent": "lappy", "text": "Lenevo", "icon": "/"},
      {"id": "hp", "parent": "lappy", "text": "HP", "icon": "/"}
  ];*/

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
        'data' : this.properties.tree
         
    }});

    //On node click
    $('#jstree').on("select_node.jstree", (e, data) => {
      const node_url = data.node.a_attr.href;
      window.location.href = node_url;
    });
    }

    //on folder change

    /*if(this.oldFolder !== this.newFolder && this.newFolder !== undefined) {
      console.log('I entered');
      $('#jstree').jstree().refresh();
    } */
  }

  //Working
  private async setSelectedFolder(folder: IFolder): Promise<void> {
    this.properties.selectedFolder = folder;
    try {
      const childFolderInfo = await this.folderService.getChildFolders(this.properties.selectedFolder.ServerRelativeUrl);
      this.properties.foldersInfo = childFolderInfo;
      const filesInfo = await this.folderService.getFilesInsideFolder(this.properties.selectedFolder.ServerRelativeUrl);
      this.properties.filesInfo = filesInfo;
      this.mappingFoldersAndFiles();
      // refresh js
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
       type: 'folder',
     });
     //mapping files
     this.properties.filesInfo.forEach(file => {
       nodeFilesAndFolders.push({
         Name: file.Name,
         id: folder.UniqueId,
         type: 'file',
         url: file.LinkingUrl
       });
     });
     this.properties.node = nodeFilesAndFolders;
     console.log(this.properties.node);
     this.mappingTree();
   });
  }

  private mappingTree(): void {
    const treeData: ITreeData[] = [];
    treeData.push({
      id: this.properties.selectedFolder.Name,
      parent: '#',
      text: this.properties.selectedFolder.Name,
      state: {
        opened: this.properties.expandAll
      }
    });

    console.log(treeData);

    this.properties.node.forEach(node => {
      treeData.push({
        id: node.id,
        parent: this.properties.selectedFolder.Name,
        text: node.Name,
        state: {
          opened: this.properties.expandAll
        },
        a_attr: {"href": node.url}
      });
    });

    console.log(treeData);

    this.properties.tree = treeData;
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
                PropertyFieldFolderPicker('selectedFolder', {
                    context: this.context,
                    label: 'Select a folder',
                    onSelect: (folder) => this.setSelectedFolder(folder),
                    rootFolder: this.properties.rootFolder,
                    selectedFolder: this.properties.selectedFolder,
                    onPropertyChange: (propertyPath: string, oldValue: IFolder, newValue: IFolder): void  => {
                       this.oldFolder = oldValue;
                       this.newFolder = newValue;
                       //this.setSelectedFolder(newValue);
                       //$('#jstree').jstree(true).refresh();
                       //$('#jstree').jstree('refresh');
                    },
                    properties: this.properties,
                    key: 'Document'
                }),
                PropertyPaneToggle('expandAll', {
                  label: 'Do you want to expand all folders?',
                  checked: this.properties.expandAll,
                }),
                PropertyPaneToggle('canCreate', {
                  label: 'Allow create option?',
                  checked: this.properties.canCreate,
                }),
                PropertyPaneToggle('canEdit', {
                  label: 'Allow edit option?',
                  checked: this.properties.canEdit,
                }),
                PropertyPaneToggle('canMove', {
                  label: 'Allow move folders/files option?',
                  checked: this.properties.canMove,
                }),
                PropertyPaneToggle('canDelete', {
                  label: 'Allow delete folders/files option?',
                  checked: this.properties.canDelete,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
