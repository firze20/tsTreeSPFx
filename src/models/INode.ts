export interface INode {
   Name: string;
   id: string;
   type: 'folder' | 'file';
   url?: string | Promise<string>;
}

