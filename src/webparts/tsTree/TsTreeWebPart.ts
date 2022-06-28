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
    //Load css js tree
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/jstree/3.2.1/themes/default/style.min.css');
    //Starting the folder service object as soon as the webpart gets loaded
    this.folderService = new FolderService(this.context);
    !this.properties.selectedFolder ? this.properties.rootFolder = await this.folderService.getRootFolder() : this.setSelectedFolder(this.properties.selectedFolder);
    return super.onInit();
  }

  public render(): void {

    if(!this.properties.selectedFolder) {
      this.domElement.innerHTML = `
       <div class="${styles.welcome}">
        <h2>JSTree <img class="${styles.jstreeIcon}" src=${require('./assets/jstree.png')} /> </h2>
        <p>Pick a folder from the web part configuration properties.</p>
        <br/>
        <p>Beta Version 0.0.1 Firze20</p>
        </div>`;
    }

    else {
      this.renderTree();
    }

  }

  private async renderTree(): Promise<void> {
    this.domElement.innerHTML = `
      <div class="${styles.divTree}">
        <div id='jstree'>
        </div>
      </div>
    
    `;

    $('#jstree').jstree({ 'core' : {
      'data' : await this.folderService.getTree(this.properties.selectedFolder, this.properties.expandAll)
       
    }});

    //On node click
    $('#jstree').on("select_node.jstree", (e, data) => {
      const node_url = data.node.a_attr.href;
      if(node_url !== '#') {
        window.open(node_url);
      }
    });

    if(this.properties.expandAll) {
      console.log(this.properties.expandAll);
      $('#jstree').jstree("open_all");
    }

    else {
      $("#jstree").jstree("close_all");
    }

  }

  //Working
  private async setSelectedFolder(folder: IFolder): Promise<void> {
    this.properties.selectedFolder = folder;
    try {
      const childFolderInfo = await this.folderService.getChildFolders(this.properties.selectedFolder.ServerRelativeUrl);
      this.properties.foldersInfo = childFolderInfo;
      const filesInfo = await this.folderService.getFilesInsideFolder(this.properties.selectedFolder.ServerRelativeUrl);
      this.properties.filesInfo = filesInfo;
      console.log(this.properties.filesInfo);
      await this.mappingFoldersAndFiles();
      // refresh js
    } catch (error) {
      console.log(error);
    }
  }

  private async mappingFoldersAndFiles(): Promise<void> {
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
      console.log(file);
       nodeFilesAndFolders.push({
         Name: file.Name,
         id: file.UniqueId,
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

                    },
                    properties: this.properties,
                    key: 'Document',
                    canCreateFolders: true,
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
