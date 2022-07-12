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

const iconFolder = require('./assets/folder-svgrepo-com.svg');
const iconFolderOpen = require('./assets/open_folder-svgrepo-com.svg');
const defaultIcon = require('./assets/file-svgrepo-com.svg');
const pdfIcon = require('./assets/pdf-svgrepo-com.svg');
const wordIcon = require('./assets/word-svgrepo-com.svg');
const excelIcon = require('./assets/excel-svgrepo-com.svg');
const imageIcon = require('./assets/image-svgrepo-com.svg');
const aspxIcon = require('./assets/aspx-svgrepo-com.svg');
const htmlIcon = require('./assets/html-svgrepo-com.svg');

export interface ITsTreeWebPartProps {
  description: string;
  rootFolder: IFolder;
  selectedFolder: IFolder;
  expandAll: boolean;
  canCreate: boolean;
  canEdit: boolean;
  canMove: boolean;
  canDelete: boolean;
}

export default class TsTreeWebPart extends BaseClientSideWebPart<ITsTreeWebPartProps> {

  private folderService: FolderService;

  //old folder and new folder

  protected async onInit(): Promise<void> {
    //Load css js tree
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/jstree/3.2.1/themes/default/style.min.css');
    //Starting the folder service object as soon as the webpart gets loaded
    this.folderService = new FolderService(this.context);
    //set root folder
    this.properties.rootFolder = await this.folderService.getRootFolder();
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
      <div class="${styles.tsTree}">
        <div id='jstree'>
        </div>
      </div>
    
    `;

    $('#jstree').jstree(
      { 
        'core' : {
          'data' : await this.folderService.getTree(this.properties.selectedFolder, this.properties.expandAll),
          'async': true,
          'check_callback': (operation, node, node_parent, node_position, more) => {
            if (operation === "move_node") {
              return node_parent.original.type === "folder"; //only allow dropping inside nodes of type 'folder'
            }
            return true;  //allow all other operations
          }
        },
        types: {
          "folder": {
            "icon" : this.properties.expandAll ? iconFolderOpen : iconFolder
          },
          "aspx": {
            "icon": aspxIcon
          },
          "pdf": {
            "icon": pdfIcon
          },
          "word": {
            "icon": wordIcon
          },
          "excel": {
            "icon": excelIcon
          },
          "image": {
            "icon": imageIcon
          },
          "html": {
            "icon": htmlIcon
          },
          "default" : {
            "icon": defaultIcon
          }

        },
        plugins: ["themes", "types", this.properties.canMove ? 'dnd' : null, this.properties.canCreate || this.properties.canEdit || this.properties.canDelete ? "contextmenu": null],
        contextmenu: {
          items: (node) => {
            const tree = $('#jstree').jstree(true);
            return {
              "Create": {
                "separator_before": false,
                "separator_after": true,
                "label": "Create",
                "action": false,
                "submenu": {
                    "File": {
                        "seperator_before": false,
                        "seperator_after": false,
                        "label": "File",
                        action: (obj) => {
                            node = tree.create_node(node, { text: 'New File', type: 'file'});
                            tree.deselect_all();
                            tree.select_node(node);
                        }
                    },
                    "Folder": {
                        "seperator_before": false,
                        "seperator_after": false,
                        "label": "Folder",
                        action: (obj) => {
                            node = tree.create_node(node, { text: 'New Folder', type: 'folder' });
                            tree.deselect_all();
                            tree.select_node(node);
                        }
                    }
                }
            },
            "Rename": {
                "separator_before": false,
                "separator_after": false,
                "label": "Rename",
                "action": (obj) => {
                    tree.edit(node);                                    
                }
            },
            "Remove": {
                "separator_before": false,
                "separator_after": false,
                "label": "Remove",
                "action": (obj) => {
                    tree.delete_node(node);
                }
            }
            };
          }
        }
      }
  );

  $('#jstree').on('open_node.jstree', (e, data) => {
    data.instance.set_icon(data.node, iconFolderOpen);
  });

  $('#jstree').on('close_node.jstree', (e, data) => {
    data.instance.set_icon(data.node, iconFolder);
  });

    //On node click
    $('#jstree').on("select_node.jstree", async (e, data) => {
      const node_url = data.node.a_attr.href;
      const node_id = data.node.id;
      const node_type = data.node.type;
      const node_name = data.node.text;
      //pdf and some files dont have a sharelink url, annonymous view needs to be generated
      if(!node_url) {
        const shareLink = await this.folderService.getShareLink(node_id);
        window.open(shareLink);
      }
      else if(node_url !== '#') {
        window.open(node_url);
      }
      else if(node_type === 'folder' && node_name !== this.properties.selectedFolder.Name) {
        const childData = await this.folderService.getChildNodes(node_id);
        //if result length
        if(childData.length > 0) {
          //because we cant add the entire array we add each object inside the array to the node in question
          childData.forEach(node => {
            if(!$('#jstree').jstree(true).get_node(node.id)) {
              $('#jstree').jstree().create_node(
                node_id, 
                node,
                'inside', 
                (result: any) => console.log(result));
            } 
          });
        }
      }
    });

    if(this.properties.expandAll) {
      $('#jstree').jstree("open_all");
    }

    else {
      $("#jstree").jstree("close_all");
    }


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
                    onSelect: (folder) =>  this.properties.selectedFolder = folder,
                    rootFolder: this.properties.rootFolder,
                    selectedFolder: this.properties.selectedFolder,
                    onPropertyChange: (propertyPath: string, oldValue: IFolder, newValue: IFolder): void  => {

                    },
                    properties: this.properties,
                    key: 'Document',
                    canCreateFolders: false, // true
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
