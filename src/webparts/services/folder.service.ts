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
import { IItemUpdateResult } from "@pnp/sp/items";

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
                            text: childFolder.Name,
                            type: 'folder',
                            state: {
                                opened: false
                            },
                            //children: true
                        }
                    );
                });
            }

            if(childFiles.length > 0) {
                childFiles.forEach(async file => {
                const extension = file.Name.substring(file.Name.lastIndexOf(".") + 1);
                    tree_data.push({
                        id: file.UniqueId,
                        parent: folder.Name,
                        text: file.Name,
                        type: this.setType(extension),
                        a_attr: { "href": file.LinkingUrl},
                        children: false
                    });
                });
            }

        } catch (error) {
            console.log(error);
        }

        return tree_data;
    }

    public async getChildNodes(folderId: string): Promise<ITreeData[]> {
        const tree_data: ITreeData[] = [];
        try {
            const folder = await this.getFolder(folderId);
            const folderRelativeUrl = folder.ServerRelativeUrl;
            const childFolders = await this.getChildFolders(folderRelativeUrl);
            const childFiles = await this.getFilesInsideFolder(folderRelativeUrl);

            if(childFolders.length > 0) {
                childFolders.forEach(childFolder => {
                    tree_data.push(
                        {
                            id: childFolder.UniqueId,
                            parent: folder.Name,
                            text: childFolder.Name,
                            type: 'folder',
                        }
                    );
                });
            }

            if(childFiles.length > 0) {
                childFiles.forEach( file => {
                    const extension = file.Name.substring(file.Name.lastIndexOf(".") + 1);
                    tree_data.push({
                        id: file.UniqueId,
                        parent: folderId,
                        text: file.Name,
                        type: this.setType(extension),
                        children: false,
                        a_attr: {"href": file.LinkingUrl}
                    });
                });
            }
        } catch (error) {
            console.log(error);
        }
        return tree_data;
    }

    /**
     * change folder Name method
     */
    public async changeName(folderId: string, newFolderName: string): Promise<IItemUpdateResult> {
         const folder = this.sp.web.getFolderById(folderId);
         const item = await folder.getItem();
         const result = await item.update({
            FileLeafRef: newFolderName
         });
         return result;
    }

    /**
     * Delete method
     */
    public async delete(folderId: string): Promise<void> {
        const folder = await this.getFolder(folderId);
        const folderrName = folder.Name;
        await this.sp.web.rootFolder.folders.getByUrl(folderrName).delete();
    }

    //returns a folder based on the id 
    private async getFolder(folderId: string): Promise<IFolder> {
        const folder = this.sp.web.getFolderById(folderId)();
        return folder;
    }

    

    private setType(fileExtension: string): string {
        switch (fileExtension) {
            case 'docx':
                return 'word';
            case 'aspx':
                return 'aspx';
            case 'doc':
                return 'word';
            case 'dotx':
                return 'word';
            case 'xlsx':
                return 'excel';
            case 'pdf':
                return 'pdf';
            case 'html':
                return 'html';
            case 'png':
                return 'image';
            case 'jpg':
                return 'image';
            case 'jpeg':
                return 'image';
            default:
                return 'default';
        }
    }

    //listItemAllFields
    public async getFolderFields(folderPath: string): Promise<ISPInstance> {
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
        //ShareLinkKind.AnnonymouseView to Organization
        const shareLink = await this.sp.web.getFolderById(fileId).getShareLink(SharingLinkKind.OrganizationView);
        return shareLink.sharingLinkInfo.Url;
    }
}