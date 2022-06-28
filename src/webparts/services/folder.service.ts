import { ISPInstance, SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/files/folder";
import "@pnp/sp/sharing";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IFolder } from "@pnp/spfx-property-controls";
import { IFileInfo } from "@pnp/sp/files/types";
import { IFolderInfo } from "@pnp/sp/folders";
import { SharingLinkKind } from "@pnp/sp/sharing";

//import { folderFromServerRelativePath } from "@pnp/sp/folders";

import {ITreeData} from '../../models';

export class FolderService {

    private context: WebPartContext; // web part context to be user with sp from pnpjs REST sharepoint api calls
    private sp: SPFI; // the sp

    constructor(_context: WebPartContext) {
        this.context = _context;
        this.sp = spfi().using(SPFx(this.context));
    }

    public async getRootFolder(): Promise<IFolder> {
        const rootFolder = await this.sp.web.rootFolder();
        return rootFolder;
    }

    public async getTree(folder: IFolder, expandNodes?: boolean): Promise<ITreeData[]> {
        const folderRelativeUrl = folder.ServerRelativeUrl;
        const tree_data: ITreeData[] = [];
        try {
            const childFolders = await this.getChildFolders(folderRelativeUrl);
            const childFiles = await this.getFilesInsideFolder(folderRelativeUrl);

            tree_data.push({
                id: folder.Name,
                parent: '#',
                text: folder.Name,
                type: 'folder',
                state: {
                    opened: expandNodes
                }
            });
            
            if(childFolders.length > 0) {
                childFolders.forEach(childFolder => {
                    tree_data.push(
                        {
                            id: childFolder.UniqueId,
                            parent: folder.Name,
                            text: childFolder.Name
                        }
                    );
                });
            }

            if(childFiles.length > 0) {
                childFiles.forEach(async file => {
                    tree_data.push({
                        id: file.UniqueId,
                        parent: folder.Name,
                        text: file.Name,
                        a_attr: { "href": file.LinkingUrl}
                    });
                });
            }

        } catch (error) {
            console.log(error);
        }

        return tree_data;
    }

    //listItemAllFields
    private async getFolderFields(folderPath: string): Promise<ISPInstance> {
        const itemFields: ISPInstance = await this.sp.web.getFolderByServerRelativePath(folderPath).listItemAllFields();
        return itemFields;
    }

    //get child folders
    private async getChildFolders(folderPath: string): Promise<IFolderInfo[]> {
        const childFolders = await this.sp.web.getFolderByServerRelativePath(folderPath).folders();
        if(childFolders.length > 0) {
            return childFolders;
        }
        else return [];
    }

    //get files inside folder
    private async getFilesInsideFolder(folderPath: string): Promise<IFileInfo[]> {
        const files = await this.sp.web.getFolderByServerRelativePath(folderPath).files();
        if(files.length > 0) {
            return files;
        }
        else return [];


    }

    //get share link for files like PDF types 
    public async getShareLink(fileId: string): Promise<string> {
        const shareLink = await this.sp.web.getFolderById(fileId).getShareLink(SharingLinkKind.AnonymousView);
        return shareLink.sharingLinkInfo.Url;
    }
}