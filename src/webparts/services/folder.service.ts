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
import { IShareLinkResponse, SharingLinkKind } from "@pnp/sp/sharing";

//import { folderFromServerRelativePath } from "@pnp/sp/folders";

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

    //listItemAllFields
    public async getFolderFiels(folderPath: string): Promise<ISPInstance> {
        const itemFields: ISPInstance = await this.sp.web.getFolderByServerRelativePath(folderPath).listItemAllFields();
        return itemFields;
    }

    //get child folders
    public async getChildFolders(folderPath: string): Promise<IFolderInfo[]> {
        const childFolders = await this.sp.web.getFolderByServerRelativePath(folderPath).folders();
        if(childFolders.length > 0) {
            return childFolders;
        }
        else return [];
    }

    //get files inside folder
    public async getFilesInsideFolder(folderPath: string): Promise<IFileInfo[]> {
        const files = await this.sp.web.getFolderByServerRelativePath(folderPath).files();
        if(files.length > 0) {
            files.forEach(async file => {
                if(!file.LinkingUrl) {
                    console.log(file.Name);
                    file.LinkingUrl = await this.getShareLink(file.ServerRelativeUrl);
                }
            });
            console.log(files);
            return files;
        }
        else return [];


    }

    //get share link for files like PDF types 
    private async getShareLink(folderPath: string): Promise<string> {
        const shareLink = await this.sp.web.getFolderByServerRelativePath(folderPath).getShareLink(SharingLinkKind.AnonymousView);
        console.log(shareLink.sharingLinkInfo.Url);
        return shareLink.sharingLinkInfo.Url;
    }
}