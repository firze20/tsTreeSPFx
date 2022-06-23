export interface ITreeData {
    "id"?: string;
    "text": string;
    "icon"?: string;
    "state"?: IState;
    "parent": string;
    "children"?: string[] | object[];
    "plugins"?: [
        "checkbox",
        "contextmenu",
	    "dnd",
	    "massload",
	    "search",
	    "sort",
	    "state",
	    "types",
	    "unique",
	    "wholerow",
	    "changed",
	    "conditionalselect"
    ];
    "url"?: string;
}

interface IState {
    "opened"?: boolean;
    "disabled"?: boolean;
    "selected"?: boolean;
}